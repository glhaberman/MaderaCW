<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: CaseAddEdit.asp                                                 '
'  Purpose: The primary data entry screen for adding new review records     '
'           and updating existing records.                                  '
'           The form is displayed when the user clicks [Enter Case Reviews] '
'           from the app main screen.                                       '
'==========================================================================='
Dim madoRs              'Generic recordset reused for various tasks.
Dim mstrPageTitle       'Title used at the top of the page.
Dim mstrAction          'The action from the post-back (add, update, etc).
Dim mlngCurrentRvwID    'Holds the record ID number of the current review.
Dim madoCmdRvw          'ADO command object for updating and getting review.
Dim madoRsRvw           'ADO recordset for updating and getting a review.
Dim mintI               'Generic loop counter.
Dim mblnPrint           'Signals a post-back after user clicks print.
Dim mintMaxCaseNumLen   'Stores the maximum length of the case number.
Dim mlngTabIndex        'Keeps track of tabindex when building controls.
Dim mblnDuplicateID     'Set True if user tries to save duplicate review.
Dim mblnDeletefail		'Set True if user tries to delete a Rereviewed review.
Dim madoCmdStf          'ADO Command to retrieve master staff list.
Dim mstrOptions         'Temp string used to build SELECT options.
Dim colPrgsElms         'Helper class for parsing program-element string.
Dim moPrgElm            'Object holding each program-element chunk.
Dim mstrReviewType      'Used to build the list of review types.
Dim mstrReviewTypeElem  'Used to build element list for each review type.
Dim mblnChangesSaved    'Set to True if page is being loaded after saving changes.
Dim mstrChangesSaved    'String to display message after saving changes.
Dim mstrElements        'Element string converted to a dictionary object client side.
Dim mstrFactors 
Dim mstrLinks
Dim mstrAllowFutureReviewDates
Dim mstrDisplayReviewGuide
Dim mdctPrograms
Dim adCmdPrg
Dim adRsPrg
Dim strHTML
Dim strGroupName
Dim mintTop, mstrRecord
Dim mintMaxFactors, moDictObj
Dim maOptions(6)
Dim mstrUserType, mdctReviewTypes
Dim mintPrgCount
    'this just a test of checking out this file
%>
<!--#include file="IncCnn.asp"-->
<!--#include file="IncValidUser.asp"-->
<%
'==============================================================================
' Server side actions:
'==============================================================================
'Instantiate the recordset that is reused for temporary results Or queries:
Set madoRs = Server.CreateObject("ADODB.Recordset")
Set mdctReviewTypes = CreateObject("Scripting.Dictionary")

'Set the page title:
mstrPageTitle = Trim(gstrTitle & " " & gstrAppName)

'Retrieve application settings needed for this page:
mintMaxCaseNumLen = GetAppSetting("MaxCaseNumberLength")
mstrAllowFutureReviewDates = GetAppSetting("AllowFutureReviewDates")
'mstrDisplayReviewGuide = GetAppSetting("DisplayReviewGuide")
mintMaxFactors = 20

mintPrgCount = 5

If InStr(gstrRoles,"[6]") > 0 Then
    mstrUserType = "W" 'Worker
End If
If InStr(gstrRoles,"[2]") > 0 Then
    mstrUserType = "S" 'Sup
End If
If InStr(gstrRoles,"[1]") > 0 Then
    mstrUserType = "O" 'Office manager
End If
If InStr(gstrRoles,"[3]") > 0 Or InStr(gstrRoles,"[4]") > 0 Then
    mstrUserType = "A" 'Admin
End If

'Determine if there is no current record, i.e. we arrived 
'here from the main menu:
mlngCurrentRvwID = ReqForm("rvwID")
If Not IsNumeric(mlngCurrentRvwID) Then
    mlngCurrentRvwID = -1
ElseIf mlngCurrentRvwID = 0 Then
    mlngCurrentRvwID = -1
End If

If mlngCurrentRvwID <> -1 Then
    'Retrieve the case to display:
    Set madoCmdRvw = GetAdoCmd("spReviewGet")
        AddParmIn madoCmdRvw, "@AliasID", adInteger, 0, glngAliasPosID
        AddParmIn madoCmdRvw, "@Admin", adBoolean, 0, gblnUserAdmin
        AddParmIn madoCmdRvw, "@QA", adBoolean, 0, gblnUserQA
        AddParmIn madoCmdRvw, "@UserID", adVarChar, 20, gstrUserID
        AddParmIn madoCmdRvw, "@rvwID", adInteger, 0, mlngCurrentRvwID
    Set madoRsRvw = Server.CreateObject("ADODB.Recordset")
    Call madoRsRvw.Open(madoCmdRvw, , adOpenForwardOnly, adLockReadOnly)
   
    If madoRsRvw.EOF Or madoRsRvw.BOF Then
        'The review not found for some reason:
        mlngCurrentRvwID = -1
    End If
End If

' Load programs
Set mdctPrograms = CreateObject("Scripting.Dictionary")
Set adRsPrg = Server.CreateObject("ADODB.Recordset")
Set adCmdPrg = GetAdoCmd("spGetProgramList")
    AddParmIn adCmdPrg, "@PrgID", adVarchar, 255, NULL
    'Call ShowCmdParms(adCmdPrg) '***DEBUG
    adRsPrg.Open adCmdPrg, , adOpenForwardOnly, adLockReadOnly
Set adCmdPrg = Nothing
Do While Not adRsPrg.EOF
    mdctPrograms.Add adRsPrg.Fields("prgID").Value, adRsPrg.Fields("prgShortTitle").Value
    adRsPrg.MoveNext 
Loop 

' Load all List Items
Set gadoCmd = GetAdoCmd("spGetListValues")
    AddParmIn gadoCmd, "@LstName", adVarChar, 50, Null
    AddParmIn gadoCmd, "@ValueID", adInteger, 0, Null
    madoRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
Set gadoCmd = Nothing

For mintI = 1 To 6
    If mintI <= 4 Then
        maOptions(mintI) = "<option value=0></option>"
    Else
        maOptions(mintI) = ""
    End If
Next

madoRs.Sort = "lstID"
Do While Not madoRs.EOF
    mintI = 0
    Select Case madoRs.Fields("lstName").Value
        Case "ElementStatus"
            mintI = 1
            maOptions(5) = maOptions(5) & madoRs.Fields("lstID").Value & "^" & madoRs("lstMemberValue").Value & "|"
        Case "TimeFrame"
            mintI = 2
        Case "ElemProgStatus"
            mintI = 3
        Case "ArrearageStatus"
            mintI = 4
            maOptions(6) = maOptions(6) & madoRs.Fields("lstID").Value & "^" & madoRs("lstMemberValue").Value & "|"
    End Select
    If mintI > 0 Then
        maOptions(mintI) = maOptions(mintI) & "<option value=" & madoRs.Fields("lstID").Value & ">" & madoRs("lstMemberValue").Value & "</option>"
    End If
    madoRs.MoveNext
Loop
madoRs.Close

' Load all Review Types
Set madoRs = Server.CreateObject("ADODB.Recordset")
Set gadoCmd = GetAdoCmd("spGetReviewTypeElms")
    AddParmIn gadoCmd, "@Programs", adVarChar, 100, Null
    AddParmIn gadoCmd, "@EffectiveDate", adDBTimeStamp, 0, Null
    madoRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
    
    
Set gadoCmd = Nothing
Do While Not madoRs.EOF
    mdctReviewTypes.Add CLng(madoRs.Fields("ReviewTypeID").Value), _
        madoRs.Fields("ReviewTypeText").Value & "^" & _
        madoRs.Fields("rteElementID").Value & "^" & _
        madoRs.Fields("rteProgramID").Value & "^" & _
        madoRs.Fields("rteStartDate").Value & "^" & _
        madoRs.Fields("rteEndDate").Value
    madoRs.MoveNext
Loop
madoRs.Close

' If user is a worker, check to see if this is the first time the user has accessed
' the record.  If it is, log it
If mstrUserType = "W" And mlngCurrentRvwID > 0 Then
    Set gadoCmd = GetAdoCmd("spReviewWrkRead")
        AddParmIn gadoCmd, "@ReviewID", adInteger, 0, mlngCurrentRvwID 
        AddParmIn gadoCmd, "@UserID", adVarchar, 20, gstrUserID
        'ShowCmdParms(gadoCmd) '***DEBUG
        gadoCmd.Execute
End If
%>

<HTML>
<HEAD>
    <META name=vs_defaultClientScript content="VBScript">
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%> <% = mstrChangesSaved %></TITLE>
    <!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
    <STYLE type="text/css">
        .ReviewTitleArea
            {
            TEXT-ALIGN:center;
            FONT-FAMILY:<%=gstrTitleFont%>;
            FONT-SIZE:<%=gstrTitleFontSmallSize%>;
            FONT-WEIGHT:bold; 
            FONT-STYLE:normal;
            COLOR:<%=gstrTitleColor%>; 
            BACKGROUND-COLOR:<%=gstrAltBackColor%>;
            BORDER-COLOR:<%=gstrBorderColor%>;
            }
        .DivTab
            {
            TOP:125; 
            WIDTH:150; 
            HEIGHT:20; 
            FONT-WEIGHT:bold;
            TEXT-ALIGN:center;
            OVERFLOW:visible;
            BORDER-BOTTOM-STYLE:none;
            BACKGROUND-COLOR:<%=gstrAltBackColor%>;
            Z-INDEX:100;
            CURSOR:default
            }
    </STYLE>
</HEAD>

<SCRIPT LANGUAGE=vbscript>
<%'==============================================================================
' Client-side script:
'==============================================================================%>
Option Explicit
Dim mblnCloseClicked    <%'Tells the window_unload event whether the close button was clicked first.%>
Dim oGuideWindow        <%'Holds reference to Guide window when it is open.%>
Dim mblnOnLoadCompleted <%'Tells Subs and Functions if they were called from within window_onload or not %>
Dim mlngTimerIDB        <%'Timer ID for building tabs %>
Dim mlngTimerIDS        <%'Timer ID for saving review %>
Dim mctlStaff           <%'Text box of staff area that is being searched for.%>
Dim mintCheckForMain    <%'Timer ID for checking if Main has been closed.%>
Dim mblnMainClosed      <%'Flag used througout page to determine if main has been closed or not.%>
Dim mdctPrograms
Dim mstrLastName
Dim mstrLastElem
Dim mstrHoldStaff
Dim mdctAudit
Dim oDictObj
Dim mintCurrentTab
Dim mblnCancelEdit
Dim mdctElmData, mdctElmComments, mdctReviewTypes
Dim maPrgs(5,1)
Dim mintArrearageID, mintDivsTabIndex
Dim mintPrgCount
<%
Response.Write "Dim maPrgTypeIDs(" & mintPrgCount & "), maElementOptions(" & mintPrgCount & ",3,1)"
%>
Sub window_onload
    Dim oOption
    Dim strElm
    Dim strTxt
    Dim intPos
    Dim intResp

    mstrLastName = "(Last Name)"
    mintCurrentTab = 1
    mblnOnLoadCompleted = False
    mblnCancelEdit = False
    mblnMainClosed = False
    mintPrgCount = <%=mintPrgCount%>
    mintCheckForMain = Window.setInterval("CheckForMain", 500) <%'Check ever 1/2 second to see if Main has been closed %>
    Call CheckForValidUser()
    Call SizeAndCenterWindow(768, 520, False)

    Set oGuideWindow = Nothing
    mblnCloseClicked = False
'<%
'    For Each oDictObj In window.opener.mdctRegions
'        Set oOption = Document.createElement("OPTION")
'        oOption.Value = oDictObj
'        oOption.Text = oDictObj
'        cboOffice.options.add oOption
'        Set oOption = Nothing
'    Next
'    For Each oDictObj In window.opener.mdctManagers
'        Set oOption = Document.createElement("OPTION")
'        oOption.Value = oDictObj
'        oOption.Text = oDictObj
'        cboManager.options.add oOption
'        Set oOption = Nothing
'    Next
' %>
    For Each oDictObj In window.opener.mdctElementIDs
        If Parse(oDictObj,"^",1) = "1" And Parse(oDictObj,"^",2) = "Arrearage" Then
            mintArrearageID = window.opener.mdctElementIDs(oDictObj)
            Exit For
        End If
    Next
    ' Load Dictionary Objects and Arrays
    Set mdctPrograms = CreateObject("Scripting.Dictionary")
    Set mdctElmData = CreateObject("Scripting.Dictionary")
    Set mdctElmComments = CreateObject("Scripting.Dictionary")
    Set mdctReviewTypes = CreateObject("Scripting.Dictionary")
    Call LoadDictionaries()
    Call LoadElementOptions("Load")
    <%'Set up the review form for initial data entry:%>
    Call LoadReviewTypes()
    Call LoadAuditDictionary()
    Call InitializeReviewEntry()
    Call divTabs_onclick(1)

    'Display the form, with review tab selected:
    divCaseBody.style.visibility = "visible"
    
    'Initialize the flag that tells when data is modified:
    Form.Changed.Value = ""
    Form.FormAction.Value = ""

    <%'Put the user in the first control:%>
    If "<%=mstrUserType%>" = "W" Then
        If txtCaseReviewID.value = "" Then
            txtReviewer.value = ""
        End If
        cmdClose.focus
    Else
        cmdAddRecord.focus
    End If

    If "<%=mstrDisplayReviewGuide%>" <> "Yes" Then
        cmdGuide.style.visibility = "hidden"
    End If

    If "<%=mstrUserType%>" = "W" And cmdChangeRecord.disabled = True Then
        Call window.opener.LoadReviewList("CASEADDEDIT")
    End If
    mblnOnLoadCompleted = True
End Sub

Sub LoadElementOptions(strAction)
    Dim intI, intJ
    Dim strRecord
    Dim dtmReviewDate, dtmElmDate, dtmElmStartDate
    
    For intI = 1 To mintPrgCount
        maElementOptions(intI,1,0) = "<option value=0></option>"
        maElementOptions(intI,2,0) = "<option value=0></option>"
        maElementOptions(intI,3,0) = "<option value=0></option>"
        maElementOptions(intI,1,1) = 0
        maElementOptions(intI,2,1) = 0
        maElementOptions(intI,3,1) = 0
    Next
    'stop
    Select Case strAction
        Case "AddRecord"
            dtmReviewDate = ""
        Case Else
            dtmReviewDate = Form.rvwDateEntered.value
    End Select
    If dtmReviewDate = "" Then dtmReviewDate = FormatDateTime(Now(),2)
    For Each oDictObj In window.opener.mdctElements
        strRecord = window.opener.mdctElements(oDictObj)
        intI = Parse(strRecord,"^",4)
        If CInt(intI) >= 50 Then intI = "6"
        intJ = Parse(strRecord,"^",5)
        dtmElmDate = Parse(strRecord,"^",3)
        If dtmElmDate = "" Then dtmElmDate = "12/31/2100"
        dtmElmStartDate = Parse(strRecord,"^",7)
        If CDate(dtmReviewDate) <= CDate(dtmElmDate) And CDate(dtmReviewDate) >= CDate(dtmElmStartDate) Then
            maElementOptions(intI, intJ, 0) = maElementOptions(intI, intJ, 0) & "<option value=" & oDictObj & ">" & Parse(strRecord,"^",1) & "</option>"
            maElementOptions(intI, intJ, 1) = maElementOptions(intI, intJ, 1) + 1
        End If
    Next
End Sub

Sub LoadDictionaries()
    <%
    For Each moDictObj In mdctPrograms
        Response.Write vbTab & "mdctPrograms.Add CLng(" & moDictObj & "), """ & mdctPrograms(moDictObj) & """" & vbCrLf
    Next
    For Each moDictObj In mdctReviewTypes
        Response.Write vbTab & "mdctReviewTypes.Add CLng(" & moDictObj & "), """ & mdctReviewTypes(moDictObj) & """" & vbCrLf
    Next
    %>
End Sub

Sub LoadReviewTypes()
    Dim oRT, oOption, intI
    Dim strRecord, dtmEnd, dtmReviewDate
    
    dtmReviewDate = Form.rvwDateEntered.value
    If dtmReviewDate = "" Then dtmReviewDate = FormatDateTime(Now(),2)
    For intI = 1 To mintPrgCount
        document.all("cboReviewType" & intI).options.length = Null
        Set oOption = Document.createElement("OPTION")
            oOption.Value = "55"
            oOption.Text = ""
            document.all("cboReviewType" & intI).options.Add oOption
        Set oOption = Nothing
    Next
    
    For Each oRT In mdctReviewTypes
        strRecord = mdctReviewTypes(oRT)
        
        dtmEnd = Parse(strRecord,"^",5)
        If dtmEnd = "" Then dtmEnd = "12/31/2100"
        
        If CDate(dtmReviewDate) >= CDate(Parse(strRecord,"^",4)) And _
            CDate(dtmReviewDate) < CDate(dtmEnd) Then
            Set oOption = Document.createElement("OPTION")
                oOption.Value = oRT
                oOption.Text = Parse(strRecord,"^",1)
                document.all("cboReviewType" & Parse(strRecord,"^",3)).options.Add oOption
            Set oOption = Nothing
        End If
    Next
End Sub

<%'If timer detects that Main has been closed, this sub will be called.  If window is
  'currently not in Edit mode, simply close the window.  If window is in Edit mode,
  'do not close window, but set the mblnMainClosed flag.  This flag will cause the
  'window to be closed at the next available opportunity. %>
Sub MainClosed()
    mblnMainClosed = True
    If cmdSaveRecord.disabled = True Then
        mblnCloseClicked = True
        window.close
    End If
End Sub

Sub LoadReviewPrograms()
    Dim intI
    Dim strRecord
    Dim strKey
    <%
    'Record Description for mdctElmData
    'Key=ProgramID^TabID^ElementID
    'Item - Delimited with "*"
    '  (1)=StatusID
    '  (2)=TimeframeID
    '  (3)=Comments
    '  (4)=FactorList - Ind factor record delimeted with "~", records delimeted with "!"
    '     (1)=FactorID
    '     (2)=FactorStatusID
    '%>
    
    mdctElmData.RemoveAll
    For intI = 1 To 1000
        strRecord = Parse(Form.ReviewElementData.value,"|",intI)
        If strRecord = "" Then Exit For
        strKey = Parse(strRecord,"^",1) & "^" & Parse(strRecord,"^",2) & "^" & Parse(strRecord,"^",3)
        
        mdctElmData.Add strKey, Parse(strRecord,"^",4)
    Next
    
    mdctElmComments.RemoveAll
    For intI = 1 To 1000
        strRecord = Parse(Form.ReviewCommentData.value,"|",intI)
        If strRecord = "" Then Exit For
        strKey = Parse(strRecord,"^",1)
        
        mdctElmComments.Add strKey, Parse(strRecord,"^",2)
    Next

    For intI = 1 To mintPrgCount
        maPrgTypeIDs(intI) = ""
    Next
    For intI = 1 To mintPrgCount
        strRecord = Parse(Form.ReviewProgramData.value, "|", intI)
        If strRecord = "" Then Exit For
        If CInt(Parse(strRecord,"^",1)) < 50 Then
            maPrgTypeIDs(Parse(strRecord,"^",1)) = Parse(strRecord,"^",2)
        End If
    Next
End Sub

Sub window_onbeforeunload
    <%'Confirm with the user before closing the browser window%>
    If Not mblnCloseClicked Then
        If Form.FormAction.value <> "" Then
            window.event.returnValue = "Closing the browser window will exit the application without saving." & space(10) & vbCrLf & "Please use the <Save> button to save your changes, then use" & space(10) & vbcrlf & "the <Close> button to return to the main menu." & space(10)
        Else
            window.event.returnValue = "Closing the browser window will exit the application." & space(10) & vbcrlf & "Please use the <Close> button to return to the main menu." & space(10)
        End If
    End If
    <%'Make sure to also close the guide window if it is still open:%>
    If Not oGuideWindow Is Nothing Then
        oGuideWindow.close
    End If
    If mblnMainClosed = False Then
        window.opener.focus
    End If
End Sub

<%'-------------------------------------------------------------------------'
'    Name: cmdPrint_onclick                                                 '
' Purpose: This event code uses a scripting dictionary object to pass all   '
'          case review data to the Print Case modal form, which handles the '
'          layout And printing.                                             '
'---------------------------------------------------------------------------'%>
Sub cmdPrint_onclick
    Dim strReturnValue
    
    document.title = "<%=Trim(gstrOrgAbbr & " " & gstrAppName)%>"
    <%'The program is designed to save a review before printing it.  If the Save
    'button is enabled, it indicates the page is in the middle of adding or 
    'modifying a review.  In that case, the "Print" keyword is appended to the 
    'value of the current FormAction, and the Save event procedure is called, 
    'causing the current review to be posted and saved to the database.  On post
    'back, the program watches for the "Print" keyword and peforms the actual
    'printing in the window onload.
    'If the review is not being edited, the print button simply causes the review
    'to print, with out posting.%>

    If Not cmdSaveRecord.disabled Then 
        If InStr(Form.FormAction.value, "Print") = 0 Then
            Form.FormAction.value = Form.FormAction.value & "Print"
        End If
        Call cmdSaveRecord_onclick
        Exit Sub
    End If
    cmdPrint.disabled = True
    <%'Open the print-preview window, passing it the review ID:%>
    strReturnValue = window.showModalDialog("PrintReview.asp?UserID=<%=gstrUserID%>&ReviewID=" & txtCaseReviewID.value, , "dialogWidth:710px;dialogHeight:520px;scrollbars:no;center:yes;border:thin;help:no;status:no")
    cmdPrint.disabled = False

    If mblnMainClosed = True Then 
        Call CloseWindow("Case Review",True,1)
        Exit Sub
    End If
    Call LoadAuditDictionary()
    Call DisplayAuditActivity()
End Sub

Sub cmdCancelEdit_onclick
    Form.FormAction.value = ""
    mblnCancelEdit = True
    If cmdCancelEdit.disabled = True Then
        Exit Sub
    End If
    Call InitializeReviewEntry()

    If "<%=mstrUserType%>" <> "W" Then
        cmdAddRecord.focus
    End If
    
    mblnCancelEdit = False
End Sub

Sub InitializeReviewEntry()
    Dim blnAllowEdit 

    <%'Initialy show the form with controls disabled:%>
    Call DisableControls(True)
    Form.SupSubWorkerDisagree.Value = "N"
    <%'Setup the initial screen to present to the user, fill the form 
    'with values from the record that is to be edited, Or default 
    'values for adding a new record:%>
    Call FillScreen
    Call DisableTabControls(True)

    If txtReviewer.value = "" Then
        <%'Fill in the review date with the current date:%>
        If Trim(txtReviewDateEntered.value) = "" Then
            txtReviewDateEntered.value = Date
            txtReviewer.value = "<%=gstrUserName%>"
        End If
    End If

    <%'Determine which buttons should be enabled:%>
    If IsNumeric(txtCaseReviewID.Value) Then
        <%'Working with an existing case review:%>
        If "<%=mstrUserType%>" = "W" Then
            cmdAddRecord.disabled = True
        Else
            cmdAddRecord.disabled = False
        End If
        cmdSaveRecord.disabled = True
        cmdChangeRecord.disabled = True
        cmdDeleteRecord.disabled = True
        <%'Decide if Change and Delete should be enabled:%>
        blnAllowEdit = False
        If chkSignature3.checked Then
            <%'Use global setting for AllowReviewEdit to determine if submitted
              'Review can be changed%>
            Select Case "<% = gstrAllowReviewEdit %>"
                Case "Manager"
                    ' Manager allows anyone with access to view the review to change it,
                    ' unless it is the same reviewer that submitted.
                    If Form.rvwUserID.value <> "<%=gstrUserID%>" Or <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
                        blnAllowEdit = True
                    End If
                Case "Any"
                    ' If user can view the review, user can change it.
                    blnAllowEdit = True
                Case "Reviewer" ' Allow original reviewer and Admins
                    If (<%=gblnUserAdmin%> Or <%=gblnUserQA%>) Or Form.rvwUserID.value <> "<%=gstrUserID%>" Then
                        blnAllowEdit = True
                    End If
                Case Else ' If not set, assume original "Admin" setting
                    ' Admin setting requires user to be an admin in order to change
                    If <%=gblnUserAdmin%> Or <%=gblnUserQA%> Then
                        blnAllowEdit = True
                    End If
            End Select
            
            If blnAllowEdit = True And "<%=mstrUserType%>" <> "W" Then
                cmdChangeRecord.disabled = False
                cmdDeleteRecord.disabled = False
            End If
        Else 'Unsubmitted Review - If user can see the review, they can edit it
            If "<%=mstrUserType%>" = "W" And Form.rvwWrkSig.value="Y" Then
                cmdChangeRecord.disabled = True
            Else
                cmdChangeRecord.disabled = False
            End If
            If "<%=mstrUserType%>" = "W" Then
                cmdDeleteRecord.disabled = True
            Else
                cmdDeleteRecord.disabled = False
            End If
        End If
        cmdCancelEdit.disabled = True
        cmdPrint.disabled = False
        cmdFindRecord.disabled = False
    Else
        <%'Working with a new case review:%>
        If "<%=mstrUserType%>" = "W" Then
            cmdAddRecord.disabled = True
        Else
            cmdAddRecord.disabled = False
        End If
        cmdSaveRecord.disabled = True
        cmdChangeRecord.disabled = True
        cmdDeleteRecord.disabled = True
        cmdCancelEdit.disabled = True
        cmdPrint.disabled = True  
        cmdFindRecord.disabled = False
    End If

    <%'Initialize the flag that tells when data is modified:%>
    Form.Changed.Value = ""
    Form.FormAction.Value = ""
End Sub

Sub cmdFindRecord_onclick
    If mblnMainClosed = True Then 
        Call CloseWindow("Case Review",True,1)
        Exit Sub
    End If
    window.opener.Form.CalledFrom.Value = "CaseAddEdit.asp"
    window.opener.Form.action = "FindCase.asp"
    Call window.opener.ManageWindows(2,"Open")
End Sub

Sub cmdAddRecord_onclick
    Form.FormAction.value = "AddRecord"
    Call ClearScreen
    Call DisableControls(False)
    Call SetButtons(True)
    txtReviewDateEntered.value = Date
    txtReviewer.value = "<%=gstrUserName%>"
    document.title = "<%=Trim(gstrOrgAbbr & " " & gstrAppName)%>"
    Call LoadElementOptions("AddRecord")
    
    txtReviewMonthYear.focus
End Sub

Sub SetButtons(blnEditing)
    cmdFindRecord.disabled = blnEditing
    If "<%=mstrUserType%>" = "W" Then
        cmdAddRecord.disabled = True
    Else
        cmdAddRecord.disabled = blnEditing
    End If
    cmdChangeRecord.disabled = blnEditing
    cmdDeleteRecord.disabled = blnEditing
    cmdSaveRecord.disabled = Not blnEditing
    cmdCancelEdit.disabled = Not blnEditing
    cmdPrint.disabled = Not blnEditing
End Sub

Sub cmdChangeRecord_onclick
    Call DisableControls(False)
    Call SetButtons(True)
    If "<%=mstrUserType%>" = "W" Then
        Call divTabs_onclick(3)
    Else
        txtReviewMonthYear.focus
    End If
    Call CheckCaseStatus()
    document.title = "<%=Trim(gstrOrgAbbr & " " & gstrAppName)%>"
    Form.FormAction.value = "ChangeRecord"
    Call LoadElementOptions("ChangeRecord")
End Sub

Sub cmdDeleteRecord_onclick
    Dim intResp
    
    intResp = MsgBox("Delete the current record?", vbQuestion + vbYesNo, "Delete")
    If intResp = vbYes Then
        Form.SaveCompleted.Value = "N"
        mlngTimerIDS = window.setInterval("CheckForCompletion",100)
        Form.rvwID.value = txtCaseReviewID.value
        Form.FormAction.value = "DeleteRecord"
        Form.action = "CaseAddEditSave.asp"
        Form.Target = "SaveFrame"
        SaveWindow.style.left = 1
        divCaseBody.style.left = -1000
        Form.Submit
    End If
    If mblnMainClosed = True Then 
        Call CloseWindow("Case Review",True,1)
        Exit Sub
    End If
End Sub

Sub CheckCorrectionDueDate()
	Dim intNowDOW 'Current day in the week
	Dim intDays   'Number of days to add
	Dim intWeeks  'Weeks to add to date
	Dim dtmDate   'Temporary date holder

	' If not checked to submit to worker, do not default due date
	If chkSignature1.checked = False Then Exit Sub
	
    If Not IsDate(txtCorrectionDue.value) Then
		' gstrWorkerRespDueDays contains both number of days AND type of days
		' First character of W=Weekdays and anything else = Days
		If Len("<% = gstrWorkerRespDueDays %>") <= 1 Then
			' Invalid code or left blank, do nothing
		Else
			intDays = Mid("<% = gstrWorkerRespDueDays %>",2,Len("<% = gstrWorkerRespDueDays %>") - 1)
			If Not IsNumeric(intDays) Then intDays = 0
			If CInt(intDays) > 0 Then
				Select Case Left("<% = gstrWorkerRespDueDays %>",1)
					Case "W","w"
						intNowDOW = WeekDay(Now(),1)
						If intDays Mod 5 = 0 Then
							' If number of days is divisible by 5, simply add weeks
							txtCorrectionDue.value = DateAdd("ww", Int(intDays/5), Now())
						Else
							' First add any whole weeks, then remaining days
							intWeeks = Int(intDays/5)
							intDays = CInt(intDays) - CInt(intWeeks*5)
							dtmDate = DateAdd("ww", intWeeks, Now())
							' If adding on remaining days puts the date past a Friday, add 2 days for weekend.
							If CInt(intDays) + CInt(intNowDOW) > 6 Then intDays = CInt(intDays) + 2
							txtCorrectionDue.value = DateAdd("d", intDays, dtmDate)
						End If
					Case Else
						txtCorrectionDue.value = DateAdd("d",intDays,Now())
				End Select
				txtCorrectionDue.value = FormatDateTime(txtCorrectionDue.value,2)
			End If
		End If
    End If
End Sub

Sub Gen_onkeydown(ctlFrom)
    If window.event.keyCode = 13 Then
        Call StaffLookUp(ctlFrom)
    End If
End Sub

Sub document_onclick()
    Select Case window.event.srcElement.id 
        Case "txtWorker", "cmdWorker"
        Case Else
            If divStaffSearch.style.left <> "-1000px" Then
                Call fraStaffSearch.LostFocus()
            End If
    End Select
End Sub

Sub txtWorkerEmpID_onfocus()
    mstrHoldStaff = txtWorkerEmpID.value
End Sub

Sub txtWorkerEmpID_onblur()
    txtWorkerEmpID.value = UCase(txtWorkerEmpID.value)
    If txtWorkerEmpID.value <> mstrHoldStaff And txtWorkerEmpID.value <> "" Then
        Call StaffLookUp(txtWorkerEmpID)
    End If
End Sub

Sub txtSupervisorEmpID_onfocus()
    mstrHoldStaff = txtSupervisorEmpID.value
End Sub

Sub txtSupervisorEmpID_onblur()
    txtSupervisorEmpID.value = UCase(txtSupervisorEmpID.value)
    If txtSupervisorEmpID.value <> mstrHoldStaff And txtSupervisorEmpID.value <> "" Then
        Call StaffLookUp(txtSupervisorEmpID)
    End If
End Sub

Sub StaffLookUp(ctlStaffID)
    Dim strType
    Dim strID

    <%'Attempt to select the reviewer for the passed ID:%>
    <%'Fill in the reviewer from ID of the logged in user:%>
    If Len(ctlStaffID.value) > 0 And ctlStaffID.value <> mstrHoldStaff Then
        Form.StaffInformation.value = ""
        divStaffSearch.style.top = 150
        Select Case ctlStaffID.ID
            Case "txtWorkerEmpID"
                Set mctlStaff = txtWorker
                strType = "txtWorker"
                strID = txtWorkerEmpID.value
                divStaffSearch.style.top = 125
                divStaffSearch.style.left = 10
            Case "txtSupervisorEmpID"
                Set mctlStaff = txtSupervisor
                strType = "txtSupervisor"
                strID = txtSupervisorEmpID.value
                divStaffSearch.style.top = 125
                divStaffSearch.style.left = 200
            Case Else
                Exit Sub
        End Select
        If strID = "?" Then strID = "%"
        fraStaffSearch.frameElement.src = "StaffSearch.asp?" & _
            "AliasID=<%=glngAliasPosID%>" & _
            "&UserAdmin=<%=gblnUserAdmin%>" & _
            "&UserQA=<%=gblnUserQA%>" & _
            "&UserID=<%=gstrUserID%>" & _
            "&Type=" & strType & _
            "&StaffName=" & _
            "&Width=300" & _
            "&StaffID=" & strID
    End If
End Sub

Sub StaffLookUpClose(strStaffInfo)
    Dim ctlIDTextBox

    divStaffSearch.style.left = -1000
    
    Set ctlIDTextBox = document.all(mctlStaff.ID & "EmpID")
    
    <%'If user clicked Cancel - revert ID field back to original value %>
    If strStaffInfo = "" Then
        <%'User cancelled search, set ID field back to original %>
        ctlIDTextBox.value = mstrHoldStaff
    ElseIf InStr(strStaffInfo,"no matches [Close]") > 0 Then
        <%'Do nothing %>
    Else
        ctlIDTextBox.value = Trim(Parse(strStaffInfo,"--",1))
        mctlStaff.value = Trim(Parse(strStaffInfo,"--",2))
    End If
    If mctlStaff.disabled = False Then mctlStaff.focus
End Sub

Sub chkSubmit_onclick()
    Dim strMsg
    Dim intResp
    
    If chkSubmit.checked Then
        If Not chkSubmitWorker.checked Then
            strMsg = ""
            strMsg = strMsg & "This review has not been submitted to the <%=gstrWkrTitle%> yet."
            strMsg = strMsg & vbCrLf & vbCrLf
            strMsg = strMsg & "Submit for reporting anyway?"
            intResp = MsgBox(strMsg, vbYesNo + vbQuestion, "Submit")
            If intResp = vbNo Then
                chkSubmit.checked = False
            End If
        End If
    End If
End Sub

Function BuildMsg(strMsg, strText)
    If strMsg = "" Then
        strMsg = strMsg & strText & space(10) & vbCrLf
    Else
        strMsg = strMsg & vbCrLf & space(4) & strText & space(10)
    End If
    
    BuildMsg = strMsg
End Function

Sub LetOptionValue(intRowID, intStatusID)
    Dim intI
    
    For intI = 0 To 3
        If intI = intStatusID - 22 Then
            document.all("optDataIntC" & intI & "R" & intRowID).checked = True
        Else
            document.all("optDataIntC" & intI & "R" & intRowID).checked = False
        End If
    Next
End Sub

Function GetOptionValue(intRowID)
    Dim intI, intStatusID
    
    intStatusID = 25
    For intI = 0 To 3
        If document.all("optDataIntC" & intI & "R" & intRowID).checked = True Then
            ' Radio buttons are Y N NA NR, which correspond to IDs 22, 23, 24, 25
            intStatusID = intI + 22
            Exit For
        End If
    Next
    
    GetOptionValue = intStatusID
End Function

Sub cmdSaveRecord_onclick
    Dim intI
    Dim strMsg
    Dim blnValidationFailed, blnErrorFound
    Dim strWorkerTitle
    Dim oElm, oPrg
    Dim blnDupCheck, blnMaxMsg
    Dim strURL
    Dim strDupFlag
    Dim intResp, intID, intRowID, intComID
    Dim blnRequired,intRowCount, blnFound
    Dim strStub, strName, strValue
    Dim strElmRecord, strFacRecord
    Dim blnSubmitted, strFocus, intTabClick, strElmMsg
    Dim strElementData, strProgramData, strCommentData, blnArrearage
    Dim dctIntegrityProgramIDs 
    Dim intProgramID, intTabID, strProgramList
    Dim strProgramName, ctlAction
    Dim strFunctionChange
    
    If divStaffSearch.style.left <> "-1000px" Then
        ' Staff search IFRAME is still visible.  Cancel the search before proceeding with Save
        Call fraStaffSearch.LostFocus()
    End If

    Set dctIntegrityProgramIDs = CreateObject("Scripting.Dictionary")
    blnDupCheck = True
    If Form.FormAction.value = "ChangeRecord" Or Form.FormAction.value = "ChangeRecordPrint" Then
        ' Only check for duplicates on edits if one of the keys has changed (Always check on an add!)
        If cboReviewClass.value = Form.rvwReviewClassID.value And _
            txtCaseNumber.value = Form.rvwCaseNumber.value And _
            txtReviewMonthYear.value = Form.rvwMonthYear.value Then
            
            blnDupCheck = False
        End If
    End If
    If blnDupCheck = True Then
        strURL = "CaseAddEditDupChk.asp?ID=" & Form.rvwID.value & "&ReviewClassID=" & cboReviewClass.value & _
            "&CaseNumber=" & txtCaseNumber.value & "&MonthYear=" & txtReviewMonthYear.value
        strDupFlag = window.showModalDialog(strURL)
        If strDupFlag = "Y" Then
            intResp = MsgBox("A Review for Case Number " & txtCaseNumber.Value & " already exists." & vbcrlf & vbcrlf & "Do you still wish to save the review?", vbQuestion + vbYesNo, "Duplicate Review")
            If intResp = vbNo Then Exit Sub
        End If
    End If
    
    strWorkerTitle = "<%=gstrWkrTitle%>"

    blnValidationFailed = False
    strMsg = ""
    strMsg = BuildMsg(strMsg, "The following items must be completed before the review can be saved:")
    
    If txtReviewer.value = "" Then
        strMsg = BuildMsg(strMsg, "Reviewer")
        If Not blnValidationFailed Then
            txtReviewer.focus
            blnValidationFailed = True
        End If
    End If

    If txtReviewMonthYear.value = vbNullString Then
        strMsg = BuildMsg(strMsg, "Review Month")
        If Not blnValidationFailed Then
            txtReviewMonthYear.focus
            blnValidationFailed = True
        End If
    End If
    
    If cboReviewClass.value = 0 Then
        strMsg = BuildMsg(strMsg, "<% = gstrReviewClassTitle %>")
        If Not blnValidationFailed Then
            cboReviewClass.focus
            blnValidationFailed = True
        End If
    End If

    'Validate the case number:
    If txtCaseNumber.value = vbNullString Then
        strMsg = BuildMsg(strMsg, "Case Number")
        If Not blnValidationFailed Then
            txtCaseNumber.focus
            blnValidationFailed = True
        End If
    End If

    If txtWorker.value = "" Or txtWorkerEmpID.value = "" Then
        strMsg = BuildMsg(strMsg, strWorkerTitle)
        If Not blnValidationFailed Then
            txtWorkerEmpID.focus
            blnValidationFailed = True
        End If
    End If
    If txtSupervisor.value = "" Or txtSupervisorEmpID.value = "" Then
        strMsg = BuildMsg(strMsg, "<%=gstrSupTitle%>")
        If Not blnValidationFailed Then
            txtSupervisorEmpID.focus
            blnValidationFailed = True
        End If
    End If

    blnErrorFound = True
    For Each oPrg In mdctPrograms
        If CInt(oPrg) < 50 Then
            If document.all("chkProgram" & oPrg).checked = True Then
                blnErrorFound = False
                Exit For
            End If
        End If
    Next
    If blnErrorFound = True Then
        strMsg = BuildMsg(strMsg, "At least 1 Program selected")
        If Not blnValidationFailed Then
            Call divTabs_onclick(1)
            chkProgram1.focus
            blnValidationFailed = True
        End If
    End If
    
    If blnValidationFailed Then
        MsgBox strMsg, vbInformation, "Save"
        Exit Sub
    End If

    blnArrearage = False
    blnSubmitted = False
    If chkSignature3.checked Or chkSignature1.checked Then blnSubmitted = True
    blnErrorFound = False
    
    intTabClick = 0
    strElmMsg = ""
    blnMaxMsg = False
    
    strElementData = ""
    If mdctElmData.Count = 0 Then
        Call DataIntegrityField_onclick(3,1)
    End If
    For Each oElm In mdctElmData
        strElmRecord = mdctElmData(oElm)
        intProgramID = CInt(Parse(oElm,"^",1))
        intTabID = CInt(Parse(oElm,"^",2))
        Select Case intTabID
            'Original system used 3 different types of elements - Type and 1 and 3 are not used here.
            Case 2 'Data Integrity
                strElementData = strElementData & oElm & "^" & strElmRecord & "|"
                If InStr(Parse(strElmRecord,"*",4),"~23") > 0 Then
                    blnErrorFound = True
                    If Not dctIntegrityProgramIDs.Exists("DI" & intProgramID) Then
                        dctIntegrityProgramIDs.Add "DI" & intProgramID, "Y"
                    End If
                Else
                    If InStr(Parse(strElmRecord,"*",4),"~22") > 0 Then
                        If Not dctIntegrityProgramIDs.Exists("DI" & intProgramID) Then
                            dctIntegrityProgramIDs.Add "DI" & intProgramID, "Y"
                        End If
                    End If
                End If
        End Select
        
        If Len(strElmMsg) > 500 Then
            strElmMsg = BuildMsg(strElmMsg, "**Additional Edits not shown**")
            blnMaxMsg = True
            Exit For
        End If
    Next

    If blnSubmitted Then
        <%'The code below enforces the rule that at least one element must be reviewed for AI.  Not applicable for Madera CW.
        'For intI = 1 To mintPrgCount
        '    If document.all("chkProgram" & intI).checked = True Then
        '        If Not dctIntegrityProgramIDs.Exists("AI" & intI) Then
        '            strElmMsg = BuildMsg(strElmMsg, mdctPrograms(CLng(intI)) & " - Action Integrity - At least 1 Action must be selected")
        '            If intTabClick = 0 Then 
        '                intTabClick = 2
        '                strFocus = "AIA"
        '            End If
        '        End If
        '    End If
        'Next
        %>
        strProgramList = "^"
        For intI = 1 To Parse(txtDataIntegrityInfo.value,"^",2)
            strValue = document.all("txtDataIntR" & intI).value
            intID = Parse(strValue,"^",1)
            If InStr(strProgramList,"^" & intID & "^") = 0 Then
                strProgramList = strProgramList & intID & "^"
            End If
            If document.all("optDataIntC3R" & intI).checked = True Then
                strName = mdctPrograms(CLng(intID))
                intID = Parse(strValue,"^",3)
                strElmMsg = BuildMsg(strElmMsg, strName & " - " & GetFactorTitle(intID) & " - Causal Factor Not Reviewed")
                If intTabClick = 0 Then 
                    intTabClick = 3
                    strFocus = "DIF~" & intI
                End If
                If Len(strElmMsg) > 500 Then
                    strElmMsg = BuildMsg(strElmMsg, "**Additional Edits not shown**")
                    blnMaxMsg = True
                    Exit For
                End If
            End If
            <%'The code below enforces the rule that comments must be entered for any causal factor marked as incorrect.  Not applicable for Madera CW.
            'If document.all("optDataIntC1R" & intI).checked = True Then
            '    strValue = document.all("txtDataIntR" & intI).value
            '    intID = Parse(strValue,"^",2)
            '    If Len(Trim(document.all("txtCommentsType2Row" & intID).value)) = 0 Then
            '        intID = Parse(strValue,"^",1)
            '        strName = mdctPrograms(CLng(intID))
            '        intID = Parse(strValue,"^",3)
            '        strElmMsg = BuildMsg(strElmMsg, strName & " - Data Integrity - " & GetFactorTitle(intID) & " - Comments")
            '        If intTabClick = 0 Then 
            '            intTabClick = 3
            '            strFocus = "DIC~" & Parse(strValue,"^",2)
            '        End If
            '        If Len(strElmMsg) > 500 Then
            '            strElmMsg = BuildMsg(strElmMsg, "**Additional Edits not shown**")
            '            blnMaxMsg = True
            '            Exit For
            '        End If
            '   End If
            'End If
            %>
        Next
        
        For intI = 2 To 11
            strValue = Parse(strProgramList,"^",intI)
            If strValue = "" Then Exit For
            If Not dctIntegrityProgramIDs.Exists("DI" & strValue) Then
                strElmMsg = BuildMsg(strElmMsg, mdctPrograms(CLng(strValue)) & " - At least 1 Causal Factor must be Yes or No")
                If intTabClick = 0 Then 
                    intTabClick = 3
                    strFocus = "DIF~"
                End If
            End If
        Next
        <%'Below code enforced the rule that if Financials program selected, the AI element Arrearage must be reviewed.  Not applicable for Madera CW.
        'If blnArrearage = False And chkProgram1.checked = True Then
        '    If Len(strElmMsg) < 500 Then
        '        strElmMsg = BuildMsg(strElmMsg, "Financials - Action Integrity - Arrearage must be reviewed")
        '        If intTabClick = 0 Then 
        '            intTabClick = 2
        '            strFocus = "AIR~"
        '        End If
        '    End If
        'End If
        %>
    End If
    
    <%'If submitting to worker or submitting for reporting, validate the review%>    
    If chkSignature3.checked Or chkSignature1.checked Then
        strMsg = ""
        If "<%=mstrUserType%>" = "W" Then
            strMsg = BuildMsg(strMsg, "The following items must be completed before the review can be saved:")
        Else
            strMsg = BuildMsg(strMsg, "The following items must be completed before the review can be submitted:")
        End If
        'Client Last/First name
        If txtClientLastName.value = vbNullString Or _
            (txtClientFirstName.value = vbNullString And txtClientFirstName.tabIndex > -1) Then
            strMsg = BuildMsg(strMsg, "Case Name")
            If Not blnValidationFailed Then
                txtClientLastName.focus
                blnValidationFailed = True
            End If
        End If
        If chkSignature2.checked And cboResponseW.value = 0 Then
            strMsg = BuildMsg(strMsg, "Worker Response")
            If Not blnValidationFailed Then
                cboResponseW.focus
                blnValidationFailed = True
            End If
        End If
        If chkSignature2.checked And cboResponseW.value = 71 And Trim(txtRvwCommentsWkr.value) = "" Then
            strMsg = BuildMsg(strMsg, "Worker Comments")
            If Not blnValidationFailed Then
                txtRvwCommentsWkr.focus
                blnValidationFailed = True
            End If
        End If
        If strElmMsg <> "" Then
            strMsg = BuildMsg(strMsg, strElmMsg)
            If intTabClick > 0 Then
                Call divTabs_onclick(intTabClick)
            End If
            strValue = Parse(strFocus,"~",2)

            Select Case Parse(strFocus,"~",1)
                Case "AIS","AIT","AIC","IGC","IGS"
                    For intI = 1 To document.all("txtProgramInfoType" & Parse(strValue,"^",2) & "Program" & Parse(strValue,"^",1)).value
                        If CInt(document.all("txtElementInfoType" & Parse(strValue,"^",2) & "Prg" & Parse(strValue,"^",1) & "Row" & intI).value) = CInt(Parse(strValue,"^",3)) Then
                            intRowID = intI
                            Exit For
                        End If
                    Next
                    strValue = "Type" & Parse(strValue,"^",2) & "Prg" & Parse(strValue,"^",1) & "Row" & intRowID
            End Select
            Select Case Parse(strFocus,"~",1)
                Case "AIS","IGS"
                    document.all("cboStatus" & strValue).focus
                Case "AIT"
                    document.all("cboTimeFrame" & strValue).focus
                Case "AIC", "IGC"
                    document.all("txtComments" & strValue).focus
                Case "AIA","DIF"
                    'document.all("cboActionType1" & strValue).focus
                Case "DIF"
                    document.all("optDataIntC0R" & Parse(strFocus,"~",2)).focus
                Case "DIC"
                    document.all("txtCommentsType2Row" & Parse(strFocus,"~",2)).focus
            End Select
            blnValidationFailed = True
        End If

        'Response and Response Due date
        If cboResponse.Value = 0 Then
            strMsg = BuildMsg(strMsg, "Worker Response Requirement")
            If Not blnValidationFailed Then
                cboResponse.focus
                blnValidationFailed = True
            End If
        End If
        If Not IsDate(txtCorrectionDue.Value) And cboResponse.Value = 235 Then
            strMsg = BuildMsg(strMsg, "Response Due Date")
            If Not blnValidationFailed Then
                txtCorrectionDue.focus
                blnValidationFailed = True
            End If
        End If
        If cboResponse.value = 235 Then
            If chkSignature3.checked Then   'Only validate when submitting to reports.
                If chkSignature2.checked = False Or cboResponseW.value = 0 Then
                    strMsg = BuildMsg(strMsg, "Worker Response:  A response must be received from the " & strWorkerTitle & " before submitting to reports.")
                    If Not blnValidationFailed Then
                        cboResponseW.focus
                        blnValidationFailed = True
                    End If
                End If
            End If
        End If
        		
        If blnValidationFailed Then
            MsgBox strMsg, vbInformation, "Save - Submit Review"
            Exit Sub
        End If
    End If

    If blnErrorFound = True Then
        Form.rvwStatusID.value = 23
    Else
        Form.rvwStatusID.value = 22
    End If
    
    strCommentData = ""
    intRowCount = Parse(txtDataIntegrityInfo.value,"^",1)
    For Each oElm In mdctElmComments
        ' Make sure screen is still visible
        blnFound = False
        For intI = 1 To intRowCount
            If oElm = document.all("lblElement" & intI).innerText Then
                blnFound = True
                Exit For
            End If
        Next
        If blnFound = True Then
            strCommentData = strCommentData & oElm & "^" & mdctElmComments(oElm) & "|"
        End If
    Next
    strProgramData = ""
    For Each oPrg In mdctPrograms
        If document.all("chkProgram" & oPrg).checked = True Then
            strProgramData = strProgramData & oPrg & "^" & document.all("cboReviewType" & oPrg).value & "|"
        End If
    Next

    If Form.rvwWrkSig.value = "Y" And chkSignature2.checked = True Then ' And (Form.rvwSubmitted.value="N" And chkSignature3.checked = True) Then
        'If review being saved was previously signed by worker, check for changes.
        If Check4Changes Or Form.ReviewProgramData.value <> strProgramData Or _
            Form.ReviewElementData.value <> strElementData Or _
            Form.ReviewCommentData.value <> strCommentData Then
            
            intResp = MsgBox("Changes have been made to this review and it has been signed by the worker.  " & vbcrlf & _
                "If you continue, your changes will be saved, but the Worker Signature will be" & vbcrlf & _
                "removed and the review will appear on the workers `Reviews That Require Attention` list." & vbCrLf & vbcrlf & _
                "Do you still wish to save the review?", vbQuestion + vbYesNo, "Signed Review")
            If intResp = vbNo Then Exit Sub
            chkSignature2.checked = False
            chkSignature3.checked = False
            cboResponseW.value = 0
        End If 
    End If
    
    <%'Save the information on the review to the HTML input form for posting:%>
    Form.DeleteCode.value = "[n]"
    strProgramList = ""
    If Form.ReviewProgramData.value <> strProgramData Then
        Form.DeleteCode.value = Form.DeleteCode.value & "[P]"
        If Form.FormAction.value = "ChangeRecord" Or Form.FormAction.value = "ChangeRecordPrint" Then
            ' Check to see if any program+review type were removed
            For intI = 1 To 100
                strElmRecord = Parse(Form.ReviewProgramData.value,"|",intI)
                If strElmRecord = "" Then Exit For
                If InStr("|" & strProgramData,"|" & strElmRecord & "|") = 0 Then
                    If Parse(strElmRecord,"^",2) <> "55" Then
                    'If Not IsReviewFull(Parse(strElmRecord,"^",2)) Then
                        strValue = mdctReviewTypes(CLng(Parse(strElmRecord,"^",2)))
                        strValue = Parse(strValue,"^",1)
                    Else
                        strValue = "Full"
                    End If
                    strProgramList = strProgramList & "Programs^" & mdctPrograms(CInt(Parse(strElmRecord,"^",1))) & " (" & strValue & ") - Present^Removed|"
                End If
            Next
            ' Check to see if any program+review type were added
            For intI = 1 To 100
                strElmRecord = Parse(strProgramData,"|",intI)
                If strElmRecord = "" Then Exit For
                If InStr("|" & Form.ReviewProgramData.value,"|" & strElmRecord & "|") = 0 Then
                    If Parse(strElmRecord,"^",2) <> "55" Then
                        strValue = mdctReviewTypes(CLng(Parse(strElmRecord,"^",2)))
                        strValue = Parse(strValue,"^",1)
                    Else
                        strValue = "Full"
                    End If
                    strProgramList = strProgramList & "Programs^" & mdctPrograms(CInt(Parse(strElmRecord,"^",1))) & " (" & strValue & ") - Not Present^Added|"
                End If
            Next
        End If
    End If
    Form.ReviewProgramData.value = strProgramData

    If Form.ReviewElementData.value <> strElementData Then
        Form.DeleteCode.value = Form.DeleteCode.value & "[E]"
    End If
    Form.ReviewElementData.value = strElementData

    If Form.ReviewCommentData.value <> strCommentData Then
        Form.DeleteCode.value = Form.DeleteCode.value & "[C]"
    End If
    Form.ReviewCommentData.value = strCommentData
    Form.SupSubWorkerDisagree.value = "N"
    If chkSignature3.checked = True And Form.rvwSubmitted.value = "N" Then
        If chkSignature2.checked = True And cboResponseW.value = 71 Then
            Form.SupSubWorkerDisagree.value = "Y"
        End If
    End If
    Call FillForm
    If strProgramList <> "" Then
        Form.UpdateString.Value = Form.UpdateString.Value & strProgramList
        Form.Changed.Value = "[Case]"
    End If

    Form.SaveCompleted.Value = "N"
    mlngTimerIDS = window.setInterval("CheckForCompletion",100)
    mblnCloseClicked = True
    Form.action = "CaseAddEditSave.asp"
    Form.Target = "SaveFrame"
    SaveWindow.style.left = 1
    divCaseBody.style.left = -1000
    Form.Submit
    
    <%' If Main has been closed, do not allow window to remain open unless Save was called from print button.%>
    If InStr(Form.FormAction.value,"Print") = 0 Then
        If mblnMainClosed = True Then
            Call CloseWindow("Case Review",True,2)
            Exit Sub
        End If
    End If
End Sub

Function CheckForCompletion()
    Dim strTitle
    Dim strFormAction
    
    If Form.SaveCompleted.value = "Y" Then
        ' Disable timer
        window.clearInterval mlngTimerIDS
        ' Page title is lost when the modal print is called.  Store title here to be reset after printing
        strTitle = window.document.title
        ' Refreshing will clear the FormAction value, so save it for later use
        strFormAction = Form.FormAction.value
        ' Update the CalledFrom flag
        Form.CalledFrom.value = "EditRecord"
        Form.FormAction.value = "EditRecord"
        
        If strFormAction = "DeleteRecord" Then
            If Form.DeleteFail.value = "N" Then
                ' Refresh screen with rvwID = 0
                Form.rvwID.value = 0
                Call ClearScreen()
            Else
                MsgBox "An Evaluation for Review Number " & Form.rvwID.value & " exists." & vbcrlf & vbcrlf , vbInformation, "Save User Login" & vbCrLf
            End If
        End If
        Call InitializeReviewEntry()
        Call LoadAuditDictionary()
        Call DisplayAuditActivity()
        Call window.opener.LoadReviewList("CASEADDEDIT")
        SaveWindow.style.left = -1000
        divCaseBody.style.left = 1
    
        If InStr(strFormAction,"Print") > 0 Then
            <%'When the user clicks the Print button on the review entry screen,
            'the form is constructed to save the record first.  After posting
            'back from the save, the form will need to finish the process by 
            'recalling the Print button event code:%>

            'Insert a call to the print button event procedure:
             Call cmdPrint_onclick()
             window.document.title = strTitle
        End If
        Form.SaveCompleted.value = ""
    End If
End Function

Sub ShowPage(blnShow)
    If blnShow Then
        fraButtons.style.visibility = "visible"
        lblDatabaseStatus.style.visibility = "hidden"
        PageBody.style.cursor = "default"
    Else
        fraButtons.style.visibility = "hidden"
        lblDatabaseStatus.style.visibility = "visible"
        PageBody.style.cursor = "wait"
    End If
End Sub

Function IsProgramSelected(strPrg)
    Dim intI
    Dim intPrg
    Dim intFindPrg

    IsProgramSelected = False
    If Not IsNumeric(strPrg) Then
        Exit Function
    End If
    
    intFindPrg = CInt(strPrg)
    For intI = 0 To cboProgram.options.length - 1
        intPrg = CInt(Trim(GetCboPrgID(intI)))
        If intFindPrg = 3 Then
            If intPrg > 4 Then
                IsProgramSelected = True
            End If
        ElseIf intFindPrg = intPrg Then
            IsProgramSelected = True
        End If
    Next
End Function

Function GetProgramTitle(strPrg)
    <%'Searches for the specificed program ID in the program 
    'master list And returns the title text for the program:%>
    Dim intI
    Dim strTitle
    
    strTitle = ""
    For intI = 0 To cboProgram.Options.length - 1
        If CStr(strPrg) = CStr(GetCboPrgID(intI)) Then
            strTitle = cboProgram.Options(intI).Text
            Exit For
        End If
    Next
    
    GetProgramTitle = strTitle
End Function

Function GetElementTitle(strElm)
    <%'Searches for the specificed element ID in the element 
    'master list and returns the short title for the element:%>
    Dim strTitle
    
    strTitle = window.opener.mdctElements.Item(CLng(strElm))
    If strTitle <> "" Then strTitle = Parse(strTitle,"^",1)
    
    GetElementTitle = strTitle
End Function

Function GetElementID(strElement, intPrgID)
    <%'Searches for the specificed element ID in the element 
    'master list and returns the short title for the element:%>
    Dim intElementID
    Dim oElm

    intElementID = 0    
    For Each oElm In window.opener.mdctElements
        If CInt(Parse(window.opener.mdctElements(oElm),"^",4)) = CInt(intPrgID) And Parse(window.opener.mdctElements(oElm),"^",1) = strElement Then
            intElementID = oElm
            Exit For
        End If
    Next
    
    GetElementID = intElementID
End Function

Function GetFactorTitle(strID)
    <%'Searches for the specificed causal factor in the factor 
    'master list and returns the causal factor text (title):%>
    Dim strTitle

    strTitle = Parse(window.opener.mdctFactors.Item(CLng(strID)),"^",1)
    
    GetFactorTitle = strTitle
End Function

Function GetFactorDescription(strID)
    <%'Searches for the specificed causal factor in the factor 
    'master list and returns the causal factor text (title):%>
    Dim strTitle

    strTitle = Parse(window.opener.mdctFactors.Item(CLng(strID)),"^",2)
    strTitle = Replace(strTitle,"[vbCrLf]",Chr(13) & Chr(10))
    strTitle = CleanTextRecordParsers(strTitle,"FromDb","All")
    
    GetFactorDescription = strTitle
End Function

Sub CleanClose()
    Call window.opener.ManageWindows(1,"Close")
End Sub

Sub cmdClose_onclick
    Dim intResp
    Dim blnClose

    If Form.FormAction.value <> "" Then
        intResp = MsgBox("You are currently editing a review" & space(10) & vbCrlf & vbCrlF & "Close the form without saving?", vbQuestion + vbYesNo, "Close Form")
        If intResp = vbYes Then
            mblnCloseClicked = True
            blnClose = True
        End If
    Else
        mblnCloseClicked = True
        blnClose = True
    End If
    
    If blnClose = True Then
        If mblnMainClosed = True Then
            <%' If main menu window is already closed, just close the window%>
            window.close
            Exit Sub
        End If
        Call window.opener.ManageWindows(1,"Close")
    End If
End Sub

Sub txtReviewMonthYear_onkeypress
    If txtReviewMonthYear.value = "(MM/YYYY)" Then
        txtReviewMonthYear.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub
Sub txtCorrectionDue_onkeypress
    If txtCorrectionDue.value = "(MM/DD/YYYY)" Then
        txtCorrectionDue.value = ""
    End If
    Call TextBoxOnKeyPress(window.event.keyCode,"D")
End Sub

Sub txtReviewMonthYear_onblur
    Dim strMonth
    Dim strYear
    Dim intPos
    Dim blnErr
    Dim blnFutureDate
    Dim dtmReviewMonth
    
    If Trim(txtReviewMonthYear.value) = "" Then
        Exit Sub
    End If
    If Trim(txtReviewMonthYear.value) = "(MM/YYYY)" Then
        txtReviewMonthYear.value = ""
        Exit Sub
    End If
    
    intPos = InStr(txtReviewMonthYear.value, "/")
    If intPos = 0 Then
        blnErr = True
    Else
        strMonth = Trim(Mid(txtReviewMonthYear.value, 1, intPos -1))
        strYear = Trim(Mid(txtReviewMonthYear.value, intPos + 1))
        If Not IsNumeric(strMonth) Or Not IsNumeric(strYear) Then
            blnErr = True
        ElseIf Len(strMonth) > 2 Or Len(strYear) > 4 Or Len(strYear) = 3 Then
            blnErr = True
        ElseIf Cint(strMonth) = 0 Or CInt(strMonth) > 12 Then
            blnErr = True
        ElseIf Len(strYear) > 2 And CInt(strYear) < 1601 Then
			blnErr = True
        Else
            If Len(strMonth) < 2 Then
                strMonth = "0" & strMonth
            End If
            If Len(strYear) <= 2 Then
                strYear = 2000 + CInt(strYear) 
            End If
            txtReviewMonthYear.value = strMonth & "/" & strYear

            blnFutureDate = False
            If "<% = mstrAllowFutureReviewDates %>" <> "Yes" Then
			    ' Do not allow Month/Year in the future
			    dtmReviewMonth = strMonth & "/01/" & strYear
			    If CDate(FormatDateTime(dtmReviewMonth,2)) > CDate(FormatDateTime(Now(),2)) Then blnFutureDate = True
            End If
        End If
    End If

    If blnErr Then
        MsgBox "Review Month Year must be in the format MM/YYYY.", vbInformation, "Case Review Entry"
        txtReviewMonthYear.focus
    End If
    If blnFutureDate Then
        MsgBox "Review Month Year cannot be after " & Right("00" & Month(Now()),2) & "/" & Year(Now()) & ".", vbInformation, "Case Review Entry"
        txtReviewMonthYear.focus
    End If
End Sub

Sub txtCaseNumber_onblur
    Dim blnErr
    
    If Trim(txtCaseNumber.value) = "" Then
        Exit Sub
    End If
    If Len(txtCaseNumber.value) > <%=mintMaxCaseNumLen%> Then
        blnErr = True
    End If
    
    If blnErr Then
        MsgBox "The Case Number must be a " & <%=mintMaxCaseNumLen%> & " character value.", vbInformation, "Case Review Entry"
        txtCaseNumber.focus
    End If
End Sub

Sub txtCorrectionDue_onblur
    If Trim(txtCorrectionDue.value) = "(MM/DD/YYYY)" Or txtCorrectionDue.value = "" Then
        txtCorrectionDue.value = ""
        Exit Sub
    End If
    If Not ValidDate(txtCorrectionDue.value) Then
        MsgBox "The Response Due Date must be a valid date - MM/DD/YYYY.", vbInformation, "Case Review Entry"
        txtCorrectionDue.focus
    Else
        If CDate(txtCorrectionDue.value) < CDate(txtReviewDateEntered.value) Then
            MsgBox "The Response Due Date can not be earlier than the review date.", vbInformation, "Case Review Entry"
            txtCorrectionDue.focus
        End If
    End If
End Sub

Sub txtClientLastName_onfocus
    If Trim(txtClientLastName.value) = "" Then
        txtClientLastName.value = mstrLastName
    End If
    txtClientLastName.select
End Sub

Sub txtClientLastName_onkeypress
    If txtClientLastName.value = mstrLastName Then
        txtClientLastName.value = ""
    End If
End Sub

Sub txtClientLastName_onblur
    If Trim(txtClientLastName.value) = "" Then
        Exit Sub
    End If
    If Trim(txtClientLastName.value) = mstrLastName Then
        txtClientLastName.value = ""
        Exit Sub
    End If
End Sub

Sub txtClientFirstName_onfocus
    If Trim(txtClientFirstName.value) = "" Then
        txtClientFirstName.value = "(First Name)"
    End If
    txtClientFirstName.select
End Sub

Sub txtClientFirstName_onkeypress
    If txtClientFirstName.value = "(First Name)" Then
        txtClientFirstName.value = ""
    End If
End Sub

Sub txtClientFirstName_onblur
    If Trim(txtClientFirstName.value) = "" Then
        Exit Sub
    End If
    If Trim(txtClientFirstName.value) = "(First Name)" Then
        txtClientFirstName.value = ""
        Exit Sub
    End If
End Sub

Sub txtReviewMonthYear_onfocus
    If Trim(txtReviewMonthYear.value) = "" Then
        txtReviewMonthYear.value = "(MM/YYYY)"
    End If
    txtReviewMonthYear.select
End Sub

Sub txtCorrectionDue_onfocus
    If Trim(txtCorrectionDue.value) = "" Then
        txtCorrectionDue.value = "(MM/DD/YYYY)"
    End If
    txtCorrectionDue.select
End Sub

<%'----------------------------------------------------------------------------
' Name:    FillScreen()
' Purpose: Fills in the review controls with the values in the FORM input 
'          fields.  It is called from the window_onload, and from the cancel
'          button code to clear and refresh the entry form.
'----------------------------------------------------------------------------%>
Sub FillScreen()
    Dim oDictObj
    Dim strRecord, strPrograms, intRowID, strName, intI
    
    Call ClearScreen()

    <%'Populate the entry form from the html FORM input fields:%>
    txtCaseReviewID.value = Form.rvwID.value
    txtReviewMonthYear.value = Form.rvwMonthYear.value
    txtReviewDateEntered.value = Form.rvwDateEntered.value
    <%'Select the reviewer and disable the reviewer field if needed:%>
    txtReviewer.value = Form.rvwReviewerName.value
	cboReviewClass.value = Form.rvwReviewClassID.Value
    txtWorkerID.value = Form.rvwWorkerID.value
    txtWorkerEmpID.value = Form.rvwWorkerEmpID.value
    txtWorker.value = Form.rvwWorkerName.value
    txtSupervisorEmpID.value = Form.rvwSupervisorEmpID.value
    txtSupervisor.value = Form.rvwSupervisorName.value
    'cboManager.value = Form.rvwManagerName.value
    'Call cboManager_onchange()
    'cboOffice.value = Form.rvwOfficeName.value
    txtClientLastName.value = Form.rvwCaseLastName.value
    txtClientFirstName.value = Form.rvwCaseFirstName.value
    txtCaseNumber.value = Form.rvwCaseNumber.value
    If Form.rvwSupSig.value = "Y" Then
        chkSignature1.checked = True
    End If
    If Form.rvwWrkSig.value = "Y" Then
        chkSignature2.checked = True
    End If
    If Form.rvwSubmitted.value = "Y" Then
        chkSignature3.checked = True
    End If
    txtRvwComments.value = CleanTextRecordParsers(Form.rvwSupComments.value,"FromDb","All")
    txtRvwCommentsWkr.value = CleanTextRecordParsers(Form.rvwWrkComments.value,"FromDb","All")
    txtCorrectionDue.value = Form.rvwResponseDueDate.value
    cboResponse.value = Form.rvwWorkerResponseID.value
    cboResponseW.value = Form.rvwWorkerSigResponseID.value

    Call LoadReviewPrograms()
    strPrograms = ""
    Dim intPrgID, intTypeID, intTypeCntr
    For Each oDictObj In mdctElmData
        intPrgID = Parse(oDictObj,"^",1)
        If CInt(intPrgID) < 50 Then
            If InStr(strPrograms,"[" & intPrgID & "]") = 0 Then
                strPrograms = strPrograms & "[" & intPrgID & "]"
                document.all("chkProgram" & intPrgID).checked = True
                If CInt(intPrgID) <> 6 Then
                    document.all("cboReviewType" & intPrgID).value = maPrgTypeIDs(intPrgID)
                End If
                Call chkProgram_onClick(intPrgID)
            End If
        End If
    Next
    <%'Fill in the overall case status:%>
    Call CheckCaseStatus()
    Call DisplayAuditActivity()
End Sub

'Fills form before save
Sub FillForm()
    Dim blnCaseChanged
    Dim intI
    Dim strDeletes
    Dim strUpdateString, strBefore, strAfter
    
    blnCaseChanged = False
    strUpdateString = ""
    If txtCaseReviewID.value <> Form.rvwID.value Then
        Form.rvwID.value = txtCaseReviewID.value
        blnCaseChanged = True
    End If
    If txtReviewMonthYear.value <> Form.rvwMonthYear.value Then
        strUpdateString = strUpdateString & "Review Month^" & Form.rvwMonthYear.value & "^" & txtReviewMonthYear.value & "|"
        Form.rvwMonthYear.value = txtReviewMonthYear.value
        blnCaseChanged = True
    End If
    If txtReviewDateEntered.value <> Form.rvwDateEntered.value Then    
        strUpdateString = strUpdateString & "Review Date^" & Form.rvwDateEntered.value & "^" & txtReviewDateEntered.value & "|"
        Form.rvwDateEntered.value = txtReviewDateEntered.value
        blnCaseChanged = True
    End If
    If txtReviewer.value <> Form.rvwReviewerName.value Then
        strUpdateString = strUpdateString & "Reviewer^" & Form.rvwReviewerName.value & "^" & txtReviewer.value & "|"
        Form.rvwReviewerName.value = txtReviewer.value
        blnCaseChanged = True
    End If
    If cboReviewClass.Value <> Form.rvwReviewClassID.Value Then
        strUpdateString = strUpdateString & "Review Class^" & Form.rvwReviewClassID.value & "^" & cboReviewClass.value & "|"
		Form.rvwReviewClassID.Value = cboReviewClass.Value
		blnCaseChanged = True
	End If
    If txtWorkerID.value <> Form.rvwWorkerID.value Then
        strUpdateString = strUpdateString & "Worker^" & Form.rvwWorkerID.value & "^" & txtWorkerID.value & "|"
        Form.rvwWorkerID.value = txtWorkerID.value
        blnCaseChanged = True
    End If
    If txtWorker.value <> Form.rvwWorkerName.value Then
        strUpdateString = strUpdateString & "Worker ID^" & Form.rvwWorkerName.value & "^" & txtWorker.value & "|"
        Form.rvwWorkerName.value = txtWorker.value
        blnCaseChanged = True
    End If
    If txtWorkerEmpID.value <> Form.rvwWorkerEmpID.value Then
        strUpdateString = strUpdateString & "Worker ID^" & Form.rvwWorkerEmpID.value & "^" & txtWorkerEmpID.value & "|"
        Form.rvwWorkerEmpID.value = txtWorkerEmpID.value
        blnCaseChanged = True
    End If
    If txtSupervisor.value <> Form.rvwSupervisorName.value Then
        strUpdateString = strUpdateString & "Supervisor Name^" & Form.rvwSupervisorName.value & "^" & txtSupervisor.value & "|"
        Form.rvwSupervisorName.value = txtSupervisor.value
        blnCaseChanged = True
    End If
    If txtSupervisorEmpID.value <> Form.rvwSupervisorEmpID.value Then
        strUpdateString = strUpdateString & "Supervisor ID^" & Form.rvwSupervisorEmpID.value & "^" & txtSupervisorEmpID.value & "|"
        Form.rvwSupervisorEmpID.value = txtSupervisorEmpID.value
        blnCaseChanged = True
    End If
    '<%
    'If cboManager.value <> Form.rvwManagerName.value Then
    '    strUpdateString = strUpdateString & "Office Manager^" & Form.rvwManagerName.value & "^" & cboManager.value & "|"
    '    Form.rvwManagerName.value = cboManager.value
    '    blnCaseChanged = True
    'End If
    'If cboOffice.value <> Form.rvwOfficeName.value Then
    '    strUpdateString = strUpdateString & "FIPs^" & Form.rvwOfficeName.value & "^" & cboOffice.value & "|"
    '    Form.rvwOfficeName.value = cboOffice.value
    '    blnCaseChanged = True
    'End If
    '%>
    If txtClientLastName.value <> Form.rvwCaseLastName.value Then
        strUpdateString = strUpdateString & "Client Last Name^" & Form.rvwCaseLastName.value & "^" & txtClientLastName.value & "|"
        Form.rvwCaseLastName.value = txtClientLastName.value  
        blnCaseChanged = True
    End If
    If txtClientFirstName.value <> Form.rvwCaseFirstName.value Then
        strUpdateString = strUpdateString & "Client First Name^" & Form.rvwCaseFirstName.value & "^" & txtClientFirstName.value & "|"
        Form.rvwCaseFirstName.value = txtClientFirstName.value
        blnCaseChanged = True
    End If
    If txtCaseNumber.value <> Form.rvwCaseNumber.value Then
        strUpdateString = strUpdateString & "Case Number^" & Form.rvwCaseNumber.value & "^" & txtCaseNumber.value & "|"
        Form.rvwCaseNumber.value = txtCaseNumber.value
        blnCaseChanged = True
    End If
    If txtCorrectionDue.value <> Form.rvwResponseDueDate.value Then
        strUpdateString = strUpdateString & "Response Due Date^" & Form.rvwResponseDueDate.value & "^" & txtCorrectionDue.value & "|"
        Form.rvwResponseDueDate.value = txtCorrectionDue.value
        blnCaseChanged = True
    End If

    'Worker Response:
    If cboResponse.value <> Form.rvwWorkerResponseID.value Then
        strAfter = cboResponse.options(cboResponse.selectedIndex).text
        strBefore = GetComboTextByID(cboResponse, Form.rvwWorkerResponseID.value)
        strUpdateString = strUpdateString & "Worker Response Requirement^" & strBefore & "^" & strAfter & "|"
        Form.rvwWorkerResponseID.value = cboResponse.value
        blnCaseChanged = True
    End If

    If cboResponseW.value <> Form.rvwWorkerSigResponseID.value Then
        strAfter = cboResponseW.options(cboResponseW.selectedIndex).text
        strBefore = GetComboTextByID(cboResponseW, Form.rvwWorkerSigResponseID.value)
        strUpdateString = strUpdateString & "Worker Response^" & strBefore & "^" & strAfter & "|"
        Form.rvwWorkerSigResponseID.value = cboResponseW.value
        blnCaseChanged = True
    End If
    
    If CleanTextRecordParsers(txtRvwComments.value,"ToDb","All") <> Form.rvwSupComments.value Then
        If Len(txtRvwComments.value) < 500 Then
            strAfter = CleanTextRecordParsers(txtRvwComments.value,"ToDb","All")
        Else
            strAfter = Left(CleanTextRecordParsers(txtRvwComments.value,"ToDb","All"),500) & "...[Truncated]"
        End If
        If Len(Form.rvwSupComments.value) < 500 Then
            strBefore = Form.rvwSupComments.value
        Else
            strAfter = Left(Form.rvwSupComments.value,500) & "...[Truncated]"
        End If
        strUpdateString = strUpdateString & "Supervisor Comments^" & strBefore & "^" & strAfter & "|"
        Form.rvwSupComments.value = CleanTextRecordParsers(txtRvwComments.value,"ToDb","All")
        blnCaseChanged = True
    End If
    If CleanTextRecordParsers(txtRvwCommentsWkr.value,"ToDb","All") <> Form.rvwWrkComments.value Then
        If Len(txtRvwCommentsWkr.value) < 500 Then
            strAfter = CleanTextRecordParsers(txtRvwCommentsWkr.value,"ToDb","All")
        Else
            strAfter = Left(CleanTextRecordParsers(txtRvwCommentsWkr.value,"ToDb","All"),500) & "...[Truncated]"
        End If
        If Len(Form.rvwWrkComments.value) < 500 Then
            strBefore = Form.rvwWrkComments.value
        Else
            strAfter = Left(Form.rvwWrkComments.value,500) & "...[Truncated]"
        End If
        strUpdateString = strUpdateString & "Worker Comments^" & strBefore & "^" & strAfter & "|"
        Form.rvwWrkComments.value = CleanTextRecordParsers(txtRvwCommentsWkr.value,"ToDb","All")
        blnCaseChanged = True
    End If
    If chkSignature1.checked Then
        If Form.rvwSupSig.value <> "Y" Then
            strUpdateString = strUpdateString & "Supervisor Signature^N^Y|"
            Form.rvwSupSig.value = "Y"
            blnCaseChanged = True
        End If
    Else
        If Form.rvwSupSig.value <> "N" Then
            strUpdateString = strUpdateString & "Supervisor Signature^Y^N|"
            Form.rvwSupSig.value = "N"
            blnCaseChanged = True
        End If
    End If

    If chkSignature2.checked Then
        If Form.rvwWrkSig.value <> "Y" Then
            strUpdateString = strUpdateString & "Worker Signature^N^Y|"
            Form.rvwWrkSig.value = "Y"
            blnCaseChanged = True
        End If
    Else
        If Form.rvwWrkSig.value <> "N" Then
            strUpdateString = strUpdateString & "Worker Signature^Y^N|"
            Form.rvwWrkSig.value = "N"
            blnCaseChanged = True
        End If
    End If

    If chkSignature3.checked Then
        If Form.rvwSubmitted.value <> "Y" Then
            strUpdateString = strUpdateString & "Submit To Reports^N^Y|"
            Form.rvwSubmitted.value = "Y"
            blnCaseChanged = True
        End If
    Else
        If Form.rvwSubmitted.value <> "N" Then
            strUpdateString = strUpdateString & "Submit To Reports^Y^N|"
            Form.rvwSubmitted.value = "N"
            blnCaseChanged = True
        End If
    End If

    If Form.Changed.Value <> "[Case]" Then
        If blnCaseChanged Then
            Form.Changed.value = "[Case]"
        End If
    End If

    Form.UpdateString.Value = strUpdateString
End Sub


Function Check4Changes()
    Dim blnCaseChanged
    
    blnCaseChanged = False
    If txtCaseReviewID.value <> Form.rvwID.value Then
        blnCaseChanged = True
    End If
    If txtReviewMonthYear.value <> Form.rvwMonthYear.value Then
        blnCaseChanged = True
    End If
    If cboReviewClass.Value <> Form.rvwReviewClassID.Value Then
		blnCaseChanged = True
	End If
    If txtWorkerID.value <> Form.rvwWorkerID.value Then
        blnCaseChanged = True
    End If
    If txtWorker.value <> Form.rvwWorkerName.value Then
        blnCaseChanged = True
    End If
    If txtWorkerEmpID.value <> Form.rvwWorkerEmpID.value Then
        blnCaseChanged = True
    End If
    If txtSupervisor.value <> Form.rvwSupervisorName.value Then
        blnCaseChanged = True
    End If
    If txtSupervisorEmpID.value <> Form.rvwSupervisorEmpID.value Then
        blnCaseChanged = True
    End If
    'If cboManager.value <> Form.rvwManagerName.value Then
    '    blnCaseChanged = True
    'End If
    'If cboOffice.value <> Form.rvwOfficeName.value Then
    '    blnCaseChanged = True
    'End If
    If txtClientLastName.value <> Form.rvwCaseLastName.value Then
        blnCaseChanged = True
    End If
    If txtClientFirstName.value <> Form.rvwCaseFirstName.value Then
        blnCaseChanged = True
    End If
    If txtCaseNumber.value <> Form.rvwCaseNumber.value Then
        blnCaseChanged = True
    End If
    If txtCorrectionDue.value <> Form.rvwResponseDueDate.value Then
        blnCaseChanged = True
    End If

    'Worker Response:
    If cboResponse.value <> Form.rvwWorkerResponseID.value Then
        blnCaseChanged = True
    End If

    If cboResponseW.value <> Form.rvwWorkerSigResponseID.value Then
        blnCaseChanged = True
    End If
    
    If CleanTextRecordParsers(txtRvwComments.value,"ToDb","All") <> Form.rvwSupComments.value Then
        blnCaseChanged = True
    End If
    If CleanTextRecordParsers(txtRvwCommentsWkr.value,"ToDb","All") <> Form.rvwWrkComments.value Then
        blnCaseChanged = True
    End If
    If chkSignature1.checked Then
        If Form.rvwSupSig.value <> "Y" Then
            blnCaseChanged = True
        End If
    Else
        If Form.rvwSupSig.value <> "N" Then
            blnCaseChanged = True
        End If
    End If

    If chkSignature2.checked Then
        If Form.rvwWrkSig.value <> "Y" Then
            blnCaseChanged = True
        End If
    Else
        If Form.rvwWrkSig.value <> "N" Then
            blnCaseChanged = True
        End If
    End If
    Check4Changes = blnCaseChanged 
End Function


<%'--------------------------------------------------------------------------
' Name:     GetCboPrgID()
' Purpose:  Parses out the program ID from the value of the program combobox.
'           If strWhichIndex is not numeric, the function defaults to taking
'           the item currently selected in the combo.  If no current item is
'           selected, the function returns a negative one.
'--------------------------------------------------------------------------%>
Function GetCboPrgID(strWhichIndex)
    Dim strID
    If Not IsNumeric(strWhichIndex) Then
        strWhichIndex = 0 'CStr(cboProgram.selectedIndex)
    End If
    If strWhichIndex = "-1" Then
        GetCboPrgId = "-1"
    Else
        'GetCboPrgID = CStr(Parse(cboProgram.options(CInt(strWhichIndex)).Value, ":", 1))
    End If
End Function

Sub ClearScreen()
    Dim intI, oPrg
    Dim strHTML
    
    mstrLastElem = ""
    <%'Move the programs from the list of selected back into the
    'master list of programs on the review summary tab:%>
    txtCaseReviewID.value = ""
    txtCaseStatus.value = ""
    txtReviewMonthYear.value = ""
    txtReviewDateEntered.value = ""
    cboReviewClass.value = 0
    txtReviewer.value = ""
    txtClientLastName.value = ""
    txtClientFirstName.value = ""
    txtCaseNumber.value = ""
    txtWorker.value = ""
    txtWorkerEmpID.value = ""
    'cboManager.value = ""
    'cboOffice.value = ""
    txtSupervisor.value = ""
    txtSupervisorEmpID.value = ""
    txtCorrectionDue.value = ""
    cboResponse.value = 0
    cboResponseW.value = 0

    divDataIntegrity.innerHTML = ""
    For intI = 1 To 5
        document.all("lblFunction" & intI).innerText = ""
        document.all("lblStatus" & intI).style.left = -1000
        If intI <=3 Then
            document.all("chkSignature" & intI).checked = False
        End If
    Next
    txtRvwComments.value = ""
    txtRvwCommentsWkr.value = ""
    mdctElmData.RemoveAll
    mdctElmComments.RemoveAll
    For intI = 1 To mintPrgCount
        maPrgTypeIDs(intI) = ""
    Next
        
    For Each oPrg In mdctPrograms
        If CInt(oPrg) < 50 Then
            document.all("chkProgram" & oPrg).checked = False
            If CInt(oPrg) <> 6 Then
                document.all("cboReviewType" & oPrg).selectedIndex = 0
            End If
        End If
    Next
    lblDIFactorDescr.innerText = ""
    
    If "<%=Request.ServerVariables("SERVER_NAME")%>" = "localhost" Then
    '=====================================================
        'for testing
        cboReviewClass.value = 261
        txtClientLastName.value = "What"
        txtClientFirstName.value = "Ever"
        txtCaseNumber.value = Timer()
        txtReviewMonthYear.value = "09/2009"

    '=====================================================
    End If
    intI = InStr(tblAudit.outerHTML,"<TBODY")
    If intI > 0 Then
        strHTML = Left(tblAudit.outerHTML,intI-1)
        tblAudit.outerHTML = strHTML & " <TBODY id=tbdAudit></TBODY></TABLE>"
    End If
End Sub

Function ReviewActionFields(strFieldName)
    ReviewActionFields = "True^<%=gstrBackColor%>"
    Select Case strFieldName
        Case "txtRvwCommentsWkr", "chkSignature2"
            If InStr("WOA","<%=mstrUserType%>") > 0 Then
                ReviewActionFields = "False^<%=gstrCtrlBackColor%>"
            End If
    End Select
    Select Case strFieldName
        Case "txtRvwComments", "chkSignature1", "chkSignature2", "chkSignature3"
            If InStr("SOA","<%=mstrUserType%>") > 0 Then
                ReviewActionFields = "False^<%=gstrCtrlBackColor%>"
            End If
    End Select
End Function

Sub SignatureOnClick(intRowID)
    If document.all("chkSignature" & intRowID).disabled = True Then Exit Sub
    document.all("chkSignature" & intRowID).checked = Not document.all("chkSignature" & intRowID).checked
    Call SignatureOnClickCtl(intRowID)
End Sub

Sub SignatureOnClickCtl(intRowID)
    Call OnChangeReviewActionTab()
End Sub

Sub cboResponse_onchange()
    Call OnChangeReviewActionTab()
End Sub

Sub OnChangeReviewActionTab()
    Dim strDisabledBackColor, strEnabledBackColor

    strDisabledBackColor = "<%=gstrBackColor%>"
    strEnabledBackColor = "<%=gstrCtrlBackColor%>"

    If InStr("WOA","<%=mstrUserType%>") > 0 Then
        If chkSignature2.checked = True Then
            cboResponseW.disabled = False
            cboResponseW.style.backgroundcolor = strEnabledBackColor
        Else
            cboResponseW.value = 0
            cboResponseW.disabled = True
            cboResponseW.style.backgroundcolor = strDisabledBackColor
        End If
    End If
    If InStr("SOA","<%=mstrUserType%>") > 0 Then
        If chkSignature1.checked = True Then
            cboResponse.disabled = False
            cboResponse.style.backgroundcolor = strEnabledBackColor
            If cboResponse.value = 235 Then 'Required
                txtCorrectionDue.disabled = False
                txtCorrectionDue.style.backgroundColor = strEnabledBackColor
                Call CheckCorrectionDueDate()
            ElseIf cboResponse.value = 232 Then 'Not Required
                txtCorrectionDue.value = ""
                txtCorrectionDue.disabled = True
                txtCorrectionDue.style.backgroundColor = strDisabledBackColor
            ElseIf cboResponse.value = 234 Then 'Not Recieved - Submitted'
                txtCorrectionDue.disabled = True
                txtCorrectionDue.style.backgroundColor = strDisabledBackColor
            End If
        Else
            txtCorrectionDue.value = ""
            cboResponse.value = 0
            cboResponse.disabled = True
            cboResponse.style.backgroundcolor = strDisabledBackColor
            txtCorrectionDue.disabled = True
            txtCorrectionDue.style.backgroundColor = strDisabledBackColor
        End If
        If chkSignature3.checked = True Then
            If chkSignature1.checked = False Then
                chkSignature3.checked = False
                MsgBox "Cannot Submit to Reports without a Supervisor Signature",vbOkOnly,"Review Signatures"
                Exit Sub
            End If
            If chkSignature1.checked = True And cboResponse.value = 235 And chkSignature2.checked = False Then
                chkSignature3.checked = False
                MsgBox "Cannot Submit to Reports without a Worker Signature" & vbCrLf & "when Worker Response Requirement is set to `Required`.",vbOkOnly,"Review Signatures"
                Exit Sub
            End If
            If chkSignature1.checked = True And cboResponse.value = 235 And chkSignature2.checked = True And cboResponseW.value = 0 Then
                chkSignature3.checked = False
                MsgBox "Cannot Submit to Reports without a Worker Response.",vbOkOnly,"Review Signatures"
                Exit Sub
            End If
        End If
    End If
End Sub

Sub DisableReviewActionTab(blnVal)
    Dim strBackColor
    Dim intI
    
    'Disable everything by default
    For intI = 1 To 3
        document.all("chkSignature" & intI).disabled = True
        document.all("chkSignature" & intI).style.backgroundColor = "<%=gstrBackColor%>"
    Next
    cboResponse.disabled = True
    cboResponse.style.backgroundcolor = "<%=gstrBackColor%>"
    cboResponseW.disabled = True
    cboResponseW.style.backgroundcolor = "<%=gstrBackColor%>"
    txtCorrectionDue.disabled = True
    txtCorrectionDue.style.backgroundColor = "<%=gstrBackColor%>"
    txtRvwComments.disabled = True
    txtRvwComments.style.backgroundColor = "<%=gstrBackColor%>"
    txtRvwCommentsWkr.disabled = True
    txtRvwCommentsWkr.style.backgroundColor = "<%=gstrBackColor%>"
    
    If blnVal = False Then
        'Enable controls based on User type and checked controls
        strBackColor = "<%=gstrCtrlBackColor%>"
        
        If InStr("SOA","<%=mstrUserType%>") > 0 Then
            chkSignature1.disabled = False
            chkSignature1.style.backgroundColor = strBackColor
            txtRvwComments.disabled = False
            txtRvwComments.style.backgroundColor = strBackColor
            If chkSignature1.checked = True Then
                cboResponse.disabled = False
                cboResponse.style.backgroundcolor = strBackColor
                If cboResponse.value = 235 Then
                    txtCorrectionDue.disabled = False
                    txtCorrectionDue.style.backgroundColor = strBackColor
                End If
            End If
            chkSignature3.disabled = False
            chkSignature3.style.backgroundColor = strBackColor
        End If
        If InStr("WOA","<%=mstrUserType%>") > 0 Then
            chkSignature2.disabled = False
            chkSignature2.style.backgroundColor = strBackColor
            txtRvwCommentsWkr.disabled = False
            txtRvwCommentsWkr.style.backgroundColor = strBackColor
            If chkSignature2.checked = True Then
                cboResponseW.disabled = False
                cboResponseW.style.backgroundcolor = strBackColor
            End If
        End If
    End If
End Sub

Sub DisableControls(blnVal)
    Dim strBackColor
    Dim intI, intJ
    Dim oPrg
    Dim intRowID
    
    If blnVal Then
        strBackColor = "<%=gstrBackColor%>"
    Else
        strBackColor = "<%=gstrCtrlBackColor%>"
    End If

    Call DisableReviewActionTab(blnVal)
    'If user is a worker, disable all remaining controls
    If "<%=mstrUserType%>" = "W" Then
        blnVal = True
        strBackColor = "<%=gstrBackColor%>"
    End If

    txtReviewMonthYear.disabled = blnVal
    txtReviewMonthYear.style.backgroundColor = strBackColor
    cboReviewClass.disabled = blnVal
    cboReviewClass.style.backgroundColor = strBackColor
    txtReviewer.disabled = blnVal
    txtReviewer.style.backgroundColor = strBackColor
    txtWorker.disabled = blnVal
    txtWorker.style.backgroundColor = strBackColor
    txtWorkerEmpID.disabled = blnVal
    txtWorkerEmpID.style.backgroundColor = strBackColor
    txtSupervisor.disabled = blnVal
    txtSupervisor.style.backgroundColor = strBackColor
    txtSupervisorEmpID.disabled = blnVal
    txtSupervisorEmpID.style.backgroundColor = strBackColor
    'cboManager.disabled = blnVal
    'cboManager.style.backgroundColor = strBackColor
    'If cboOffice.options.length > 2 Then
    '    cboOffice.disabled = blnVal
    'Else
    '    cboOffice.disabled = True
    'End If
    'cboOffice.style.backgroundColor = strBackColor
    txtClientLastName.disabled = blnVal
    txtClientLastName.style.backgroundColor = strBackColor
    txtClientFirstName.disabled = blnVal
    txtClientFirstName.style.backgroundColor = strBackColor
    txtCaseNumber.disabled = blnVal
    txtCaseNumber.style.backgroundColor = strBackColor
    Call DisableTabControls(blnVal)
End Sub

Sub DisableTabControls(blnVal)
    Dim strBackColor
    Dim intI, intJ, intK
    Dim oPrg
    Dim intRowID
    Dim strRecord, strFactor
    
    If blnVal Then
        strBackColor = "<%=gstrBackColor%>"
    Else
        strBackColor = "<%=gstrCtrlBackColor%>"
    End If
    intJ = 0
    'For Each oPrg In mdctPrograms
    For oPrg = 1 To mintPrgCount
        document.all("chkProgram" & oPrg).disabled = blnVal
        If document.all("chkProgram" & oPrg).checked = True Then
            document.all("cboReviewType" & oPrg).disabled = blnVal
            document.all("cboReviewType" & oPrg).style.backgroundColor = strBackColor
            intJ = intJ + 1
        Else
            document.all("cboReviewType" & oPrg).disabled = True
            document.all("cboReviewType" & oPrg).style.backgroundColor = "<%=gstrBackColor%>"
        End If
    Next
    
    If intJ > 0 Then
        ' Data Integrity
        For intI = 1 To Parse(txtDataIntegrityInfo.value,"^",1)
            document.all("txtCommentsType2Row" & intI).disabled = blnVal
            For Each oPrg In mdctPrograms
                If InStr("F" & document.all("txtDIProgramIDList" & intI).value,"F" & oPrg & "F") > 0 Then
                    document.all("cmdScreenNAPrg" & oPrg & "R" & intI).disabled = blnVal
                End If
            Next
        Next
        For intI = 1 To Parse(txtDataIntegrityInfo.value,"^",2)
            For intJ = 0 To 3
                document.all("optDataIntC" & intJ & "R" & intI).disabled = blnVal
            Next
        Next
    End If
End Sub

Function CleanText(strText, strDir)
    <%'This function is used to replace quote and double-quote characters with 
    'tokens when sending to the database, and replace the tokens with the
    'correct characters when retrieving from the database. The tokens used 
    'are {TAB}#sq# for single-quote
    '    {TAB}#dq# for double-quote%>

    If IsNull(strText) Then
        CleanText = ""
    Else 'Apostrophe Or double quotes:
        If strDir = "FromDb" Then
            strText = Replace(strText, Chr(9) & "#sq#", "'")
            strText = Replace(strText, Chr(9) & "#dq#", """")
        ElseIf strDir = "ToDb" Then
            strText = Replace(strText, "'", Chr(9) & "#sq#")
            strText = Replace(strText, """", Chr(9) & "#dq#")
        End If
        CleanText = strText
    End If
End Function

Function CleanTextRecordParsers(strText, strDir, strType)
    <%'This function is used to replace carrot and bar characters with 
    'tokens when sending to the database, and replace the tokens with the
    'correct characters when retrieving from the database. The tokens used 
    'are {TAB}#ca# for carrot
    '    {TAB}#ba# for bar%>

    If IsNull(strText) Then
        CleanTextRecordParsers = ""
    Else 'Apostrophe Or double quotes:
        If strDir = "FromDb" Then
            strText = Replace(strText, Chr(9) & "#ca#", "^")
            strText = Replace(strText, Chr(9) & "#ba#", "|")
            strText = Replace(strText, Chr(9) & "#as#", "*")
            strText = Replace(strText, Chr(9) & "#ex#", "!")
            If strType = "All" Then
                strText = Replace(strText, Chr(9) & "#sq#", "'")
                strText = Replace(strText, Chr(9) & "#dq#", """")
            End If
        ElseIf strDir = "ToDb" Then
            strText = Replace(strText, "^", Chr(9) & "#ca#")
            strText = Replace(strText, "|", Chr(9) & "#ba#")
            strText = Replace(strText, "*", Chr(9) & "#as#")
            strText = Replace(strText, "!", Chr(9) & "#ex#")
            If strType = "All" Then
                strText = Replace(strText, "'", Chr(9) & "#sq#")
                strText = Replace(strText, """", Chr(9) & "#dq#")
            End If
        End If
        CleanTextRecordParsers = strText
    End If
End Function

<%'--------------------------------------------------------------------------
' Name:     ReturnNumeric()
' Purpose:  This function loops through the string value that was passed in
'   and removes any non-numeric characters, returning only the numeric parts
'   of the string.
'--------------------------------------------------------------------------%>
Function ReturnNumeric(strVal)
    Dim intPos
    Dim strTmp
    
    For intPos = 1 To Len(strVal)
        If IsNumeric(Mid(strVal, intPos, 1)) Then
            strTmp = strTmp & Mid(strVal, intPos, 1)
        End If
    Next
    
    ReturnNumeric = strTmp
End Function

Sub document_onkeydown
    <%'Treats the press of the <Escape> key as if the <Cancel> button were clicked.%>
    If window.event.keyCode = 27 Then
        If cmdCancelEdit.disabled = False Then
            Call cmdCancelEdit_onclick
        End If
    End If
End Sub

Function LastElement(strAction, intProgramID, intElementIndex)
    Dim strKey
    Dim intI, intJ
    Dim strFind
    
    strKey = "[" & intProgramID & "^"
    intI = InStr(mstrLastElem, strKey)
    If strAction = "Get" Then
        If intI > 0 Then
            strFind = Mid(mstrLastElem, intI + Len(strKey))
            intJ = Parse(strFind,"]",1)
            If Not IsNumeric(intJ) Then intJ = 0
        Else
            intJ = 0
	    End If
	    LastElement = intJ
	ElseIf strAction = "Set" Then
	    If intI > 0 Then
            strFind = Mid(mstrLastElem, intI + Len(strKey))
            intJ = Parse(strFind,"]",1)
            mstrLastElem = Replace(mstrLastElem, strKey & intJ, strKey & intElementIndex)
        Else
            mstrLastElem = mstrLastElem & strKey & intElementIndex & "]"
	    End If
	    LastElement = -1
    End If
End Function

Sub ExpandReviewComments(ctlButton)
    Dim ctlLabel, ctlTextBox, ctlOtherLabel, ctlOtherTextBox, ctlOtherCmd
    Dim intTop
    
    If ctlButton.ID = "cmdExpandRvwComments" Then
        Set ctlLabel = lblRvwComments
        Set ctlTextBox = txtRvwComments
        Set ctlOtherLabel = lblRvwCommentsWkr
        Set ctlOtherTextBox = txtRvwCommentsWkr
        Set ctlOtherCmd = cmdExpandRvwCommentsWkr
        intTop = 7
    Else
        Set ctlLabel = lblRvwCommentsWkr
        Set ctlTextBox = txtRvwCommentsWkr
        Set ctlOtherLabel = lblRvwComments
        Set ctlOtherTextBox = txtRvwComments
        Set ctlOtherCmd = cmdExpandRvwComments
        intTop = 74
    End If
    
    If ctlButton.innerText = "/\" Then
        ctlButton.innerText = "\/"
        ctlButton.Title = "Condense Review Comments Field"
        ctlButton.style.top = 2
        ctlLabel.style.left = 5
        ctlLabel.style.top = 7
        ctlTextBox.style.left = 5
        ctlTextBox.style.top = 22
        ctlTextBox.style.height = 110
        ctlTextBox.style.width = 730
        ctlOtherLabel.style.left = -1000
        ctlOtherTextBox.style.left = -1000
        ctlOtherCmd.style.left = -1000
        divCheckBoxes.style.left = -1000
    Else
        ctlButton.innerText = "/\"
        ctlButton.Title = "Expand Review Comments Field"
        ctlButton.style.top = intTop - 5 '2
        ctlLabel.style.left = 320
        ctlLabel.style.top = intTop '7
        ctlTextBox.style.left = 320
        ctlTextBox.style.top = intTop + 15 '22
        ctlTextBox.style.height = 45
        ctlTextBox.style.width = 410
        ctlOtherLabel.style.left = 320
        ctlOtherTextBox.style.left = 320
        ctlOtherCmd.style.left = 711
        divCheckBoxes.style.left = 5
    End If
End Sub

Sub divTabs_onkeydown(intTab)
    If window.event.keyCode = 32 Then
        If Document.all("divTab" & intTab).style.left = "-5000px" Then
            Call divTabs_onclick(intTab)
        End If
    End If
End Sub

Sub divTabs_onclick(intTab)
    Dim intI, intProgramID            <%'Loop through programs collection.%>
    
    If intTab = 3 Then
        Call CheckCaseStatus()
    End If
    If intTab = 2 Then
        For intProgramID = 1 To 5
            If document.all("chkProgram" & intProgramID).checked = True Then
                If CStr(document.all("cboReviewType" & intProgramID).value) = "55" Then
                    MsgBox "Please select a review type."
                    Exit Sub
                End If
            End If
        Next
    End If
     
    <%'Exit if the tab click is somehow called on a tab that's not visible:%>
    If Document.all("divTab" & intTab).style.left = "-5000pt" Then
        Exit Sub
    End If
    'If divFunctionsLoading.style.posLeft > 0 Then
    If Form.TabsDisabled.value = "True" Then
        Exit Sub
    End If
    <%'Make the contents of the selected tab visible, and hide the contents
    'of the other tabs.%>
    For intI = 1 To 3
        If intI = intTab Then
            Document.all("divTabButton" & intI).style.borderBottomStyle = "none"
            Document.all("divTabButton" & intI).style.fontWeight = "bold"
            Document.all("divTab" & intI).style.left = 0
            Document.all("divTab" & intI).style.visibility = "visible"
        Else
            Document.all("divTabButton" & intI).style.borderBottomStyle = "solid"
            Document.all("divTabButton" & intI).style.fontWeight = "normal"
            Document.all("divTab" & intI).style.left = -5000
            Document.all("divTab" & intI).style.visibility = "hidden"
        End If
    Next 
End Sub

'strKey = "Type" & intElementTypeID & "Prg" & intProgramID & "Row" & intRowID
Sub ElementStatus_onchange(intTabID, intRowID, intPrgID)
    Dim strKey, strDecStatus
    Dim intTimeFrameID
    Dim blnContinue

    strKey = "Type" & intTabID & "Prg" & intPrgID & "Row" & intRowID
    strDecStatus = GetDecisionStatus(strKey)
    If document.all("txtElementInfo" & strKey).value = "0" Then
        If intTabID = 1 Then
            MsgBox "Action Status cannot be set until an Action is selected.", vbOkOnly,"Action Integrity"
        Else
            MsgBox "Answer cannot be selected until a Question is selected.", vbOkOnly,"Information Gathering"
        End If
        document.all("cboStatus" & strKey).value = "0"
        Exit Sub
    End If
    blnContinue = True
    If intTabID = 1 Then
        If strDecStatus = "0" And document.all("cboStatus" & strKey).value <> "0" Then
            MsgBox "Action Status cannot be set until all Decisions are completed.", vbOkOnly,"Action Integrity"
            document.all("cboStatus" & strKey).value = "0"
            blnContinue = False
        End If
        If document.all("cboStatus" & strKey).value = "30" And blnContinue = True Then
            If strDecStatus <> "22" And strDecStatus <> "-1" Then
                MsgBox "Action Status cannot be Correct unless all Decisions are Yes or NA.", vbOkOnly,"Action Integrity"
                document.all("cboStatus" & strKey).value = "0"
                blnContinue = False
            Else
                document.all("cboTimeFrame" & strKey).disabled = False
            End If
        ElseIf document.all("cboStatus" & strKey).value <> "0" And blnContinue = True Then
            If strDecStatus = "22" Then
                MsgBox "Action Status must be Correct if all Decisions are Yes or NA.", vbOkOnly,"Action Integrity"
                document.all("cboStatus" & strKey).value = "30"
                document.all("cboTimeFrame" & strKey).disabled = False
                blnContinue = False
            Else
                document.all("cboTimeFrame" & strKey).disabled = True
                document.all("cboTimeFrame" & strKey).value = "0"
            End If
        ElseIf blnContinue = True Then
            document.all("cboTimeFrame" & strKey).disabled = True
            document.all("cboTimeFrame" & strKey).value = "0"
        End If
        intTimeFrameID = document.all("cboTimeFrame" & strKey).value
    Else
        intTimeFrameID = 0
    End If
    Call mdctElmData_UpdateElement(intPrgID & "^" & intTabID & "^" & document.all("txtElementInfo" & strKey).value, _
        document.all("cboStatus" & strKey).value, _
        intTimeFrameID, _
        document.all("txtComments" & strKey).value)
End Sub

Sub ElementTimeframe_onchange(intTabID, intRowID, intPrgID)
    Dim strKey
    
    strKey = "Type" & intTabID & "Prg" & intPrgID & "Row" & intRowID
    Call mdctElmData_UpdateElement(intPrgID & "^" & intTabID & "^" & document.all("txtElementInfo" & strKey).value, _
        document.all("cboStatus" & strKey).value, _
        document.all("cboTimeFrame" & strKey).value, _
        document.all("txtComments" & strKey).value)
End Sub


Sub ElementCommentDI_onblur(intRowID)
    Dim strScreenName
    
    strScreenName = document.all("lblElement" & intRowID).innerText
    
    If mdctElmComments.Exists(strScreenName) Then
        mdctElmComments(strScreenName) = CleanTextRecordParsers(document.all("txtCommentsType2Row" & intRowID).value,"ToDb","All")
    Else
        mdctElmComments.Add strScreenName, CleanTextRecordParsers(document.all("txtCommentsType2Row" & intRowID).value,"ToDb","All")
    End If
End Sub

Sub ElementComment_onblur(intTabID,intRowID, intPrgID)
    Dim strKey, intTimeFrameID

    strKey = "Type" & intTabID & "Prg" & intPrgID & "Row" & intRowID

    If CLng(document.all("cboAction" & strKey).value) = 0 Then Exit Sub
    
    If intTabID = 1 Then
        intTimeFrameID = document.all("cboTimeFrame" & strKey).value
    Else
        intTimeFrameID = 0
    End If
    Call mdctElmData_UpdateElement(intPrgID & "^" & intTabID & "^" & document.all("txtElementInfo" & strKey).value, _
        document.all("cboStatus" & strKey).value, _
        intTimeFrameID, _
        document.all("txtComments" & strKey).value)
End Sub

Function GetDecisionStatus(strKey)
    Dim strRecord, strFactorList
    Dim intI, strStatus
    
    strRecord = window.opener.mdctElements(CLng(document.all("txtElementInfo" & strKey).value))
    strFactorList = RemoveInactiveFactors(Parse(strRecord,"^",8))
    
    If strFactorList = "0*" Then
        ' No decisions
        GetDecisionStatus = "-1"
        Exit Function
    End If

    strStatus = "0"
    For intI = 1 To 100
        strRecord = Parse(strFactorList,"*",intI)
        If strRecord = "" Then Exit For
        Select Case document.all("cboStatus" & strKey & "F" & strRecord).value
            Case "0" ' Blank
                strStatus = "0"
                Exit For
            Case "22" 'Yes
                If strStatus = "0" Or strStatus = "24" Then
                    strStatus = "22"
                End If
            Case "23" 'No
                strStatus = "23"
            Case "24" 'NA
                If strStatus = "0" Then
                    strStatus = "24"
                End If
        End Select
    Next
    
    GetDecisionStatus = strStatus
End Function

Sub FactorStatus_onchange(intTabID, intRowID, intPrgID, intFactorID)
    Dim blnElementChanged
    Call mdctElmData_UpdateFactor(intPrgID & "^" & intTabID & "^" & document.all("txtElementInfoType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value, _
        intFactorID, document.all("cboStatusType" & intTabID & "Prg" & intPrgID & "Row" & intRowID & "F" & intFactorID).value)
    
    blnElementChanged = False
    Select Case GetDecisionStatus("Type" & intTabID & "Prg" & intPrgID & "Row" & intRowID)
        Case "22"
            If document.all("cboStatusType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value <> "30" Then
                blnElementChanged = True
            End If
            document.all("cboStatusType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value = "30"
            document.all("cboTimeFrameType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).disabled = False
        Case Else
            If document.all("cboStatusType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value <> "0" Then
                blnElementChanged = True
            End If
            document.all("cboStatusType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value = "0"
            document.all("cboTimeFrameType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).value = "0"
            document.all("cboTimeFrameType" & intTabID & "Prg" & intPrgID & "Row" & intRowID).disabled = True
    End Select
    If blnElementChanged = True Then
        Call ElementStatus_onchange(intTabID, intRowID, intPrgID)
    End If
End Sub

Sub DataIntegrityField_onclick(intOptionID, intControlID) ', intProgramID, intElementID, intFactorID, intOptionID)
    Dim intI, strRecord
    Dim strKey, intProgramID, intElementID, intFactorID, intRowID
    
    strRecord = document.all("txtDataIntR" & intControlID).value
    intProgramID = Parse(strRecord,"^",1)
    intRowID = Parse(strRecord,"^",2)
    intFactorID = Parse(strRecord,"^",3)
    intElementID = window.opener.mdctElementIDs(Parse(strRecord,"^",1) & "^" & document.all("lblElement" & Parse(strRecord,"^",2)).innerText)
    
    For intI = 0 To 3
        If intI = intOptionID Then
            document.all("optDataIntC" & intI & "R" & intControlID).checked = True
        Else
            document.all("optDataIntC" & intI & "R" & intControlID).checked = False
        End If
    Next
    
    Call mdctElmData_UpdateFactor(intProgramID & "^2^" & intElementID, _
        intFactorID, CInt(intOptionID) + 22)
End Sub

Sub cboReviewType_onClick(intRowID)
    Dim intI
    Dim strRecord
    Dim blnChanged, intResp
    Dim oElm
    
    If CInt(maPrgTypeIDs(intRowID)) <> CInt(document.all("cboReviewType" & intRowID).value) Then
        'If CInt(document.all("cboReviewType" & intRowID).value) <> 55 Then
            blnChanged = False
            For Each oElm In mdctElmData
                If CInt(Parse(oElm,"^",1)) = CInt(intRowID) Then
                    blnChanged = True
                    Exit For
                End If
            Next
            If blnChanged = True Then
                intResp = MsgBox("Changing the Review Type will result in all Elements for this program being cleared." & vbCrLf & vbCrLf & "Are you sure you want to change the Review Type?",vbYesNo,"Case Review Entry")
                If intResp = vbNo Then
                    document.all("cboReviewType" & intRowID).value = maPrgTypeIDs(intRowID)
                    Exit Sub
                End If
            End If
        'End If
    End If
    
    maPrgTypeIDs(intRowID) = document.all("cboReviewType" & intRowID).value
    'If CInt(document.all("cboReviewType" & intRowID).value) <> 55 Then
        For Each oElm In mdctElmData
            If CInt(Parse(oElm,"^",1)) = CInt(intRowID) Then
                mdctElmData.Remove oElm
            End If
        Next
    'End If
    Call chkProgram_onClick(intRowID)
End Sub

Sub RemoveFunctionFromDictionary(strRowID)
    Dim oElm, oPrg
    Dim intRowID, strProgramName, intProgramID
    
    intProgramID = -1
    intRowID = Parse(strRowID,"^",2)
    If Parse(strRowID,"^",1) = "F" Then 
        'Function ID passed in
        intProgramID = intRowID
    Else
        'Enfor Rem Action ID (element ID) passed in. Need to convert to element name and then
        'determine program ID from the name
        strProgramName = Parse(window.opener.mdctElements(CLng(intRowID)),"^",1)
        For Each oPrg In mdctPrograms
            If mdctPrograms(oPrg) = strProgramName Then
                intProgramID = oPrg
                Exit For
            End If
        Next
    End If

    If CInt(intProgramID) > 0 Then
        For Each oElm In mdctElmData
            If CInt(Parse(oElm,"^",1)) = CInt(intProgramID) Or (CInt(intProgramID)=6 And CInt(Parse(oElm,"^",1))>=50) Then
                mdctElmData.Remove oElm
            End If
        Next
    End If
End Sub

Sub chkProgram_onClick(intRowID)
    If document.all("chkProgram" & intRowID).checked = True Then
        maPrgTypeIDs(intRowID) = "0"
    Else
        maPrgTypeIDs(intRowID) = ""
        document.all("cboReviewType" & intRowID).selectedIndex = 0
        Call RemoveFunctionFromDictionary("F^" & intRowID)
    End If
    
    If mblnOnLoadCompleted = False Or mblnCancelEdit = True Or Form.SaveCompleted.Value = "Y" Then
        Call BuildTabs(intRowID)
    Else
        divFunctions.style.left = -1000
        divFunctionsLoading.style.left =10
        document.all("lblProgramD" & intRowID).innerText = document.all("lblProgram" & intRowID).innerText & " (Building...)"
        Form.TabBuildCompleted.value = "S^" & intRowID
        Form.TabsDisabled.value = "True"
        mlngTimerIDB = window.setInterval("CheckForBuildCompletion",100)
    End If
    lblDIFactorDescr.innerText = ""
End Sub

Function CheckForBuildCompletion()
    Dim intRowID
    If Left(Form.TabBuildCompleted.value,1) = "Y" Then
        window.clearInterval mlngTimerIDB
        intRowID = Parse(Form.TabBuildCompleted.value,"^",2)
        divFunctions.style.left = 10
        divFunctionsLoading.style.left =-1000
        document.all("lblProgramD" & intRowID).innerText = document.all("lblProgram" & intRowID).innerText
        Form.TabsDisabled.value = "False"
    ElseIf Left(Form.TabBuildCompleted.value,1) = "S" Then
        intRowID = Parse(Form.TabBuildCompleted.value,"^",2)
        Call BuildTabs(intRowID)
    ElseIf Left(Form.TabBuildCompleted.value,1) = "A" Then
        Call BuildDataIntegrityTab(Parse(Form.TabBuildCompleted.value,"^",2))
    ElseIf Left(Form.TabBuildCompleted.value,1) = "Z" Then
        window.clearInterval mlngTimerIDB
        Form.TabsDisabled.value = "False"
        document.all("cboAction" & Parse(Form.TabBuildCompleted.value,"^",2)).disabled = False
    End If
End Function

Sub BuildTabs(intRowID)
    Dim strBuild
    Dim intI, intProgramID
    Dim strRecord
    Dim intTop
    Dim oDictObj
    Dim blnPrintName, blnFound
    
    ' Because the controls are all built by changing the innerHTML, GetTabIndex will
    ' not incriment properly.  Set a local variable to GetTabIndex and use that to increment.
    mintDivsTabIndex = 100
    Form.TabBuildCompleted.value = "N^" & intRowID
    ' Data Integrity Tab
    Call BuildDataIntegrityTab("BuildTab")
    Call DisableTabControls(chkProgram1.disabled)
    Form.TabBuildCompleted.value = "Y^" & intRowID
End Sub 

Function IsReviewFull(intReviewTypeID)
    If CLng(intReviewTypeID) > 40 And CLng(intReviewTypeID) <=55 Then
        IsReviewFull = True
    Else
        IsReviewFull = False
    End If
End Function

Sub BuildDataIntegrityTab(strWhoCalled)
    Dim dctBuild
    Dim dctPrgs, oDictObj, oElm
    Dim strPrograms, strScreenName, strBuild
    Dim intI, intTop, intJ, intK, intProgramID
    Dim strFactorList, strFactor, intFactorID
    Dim strHoldScreen, strProgramList, strRecord, blnIncludeFull
    Dim blnProgramSelected
    Dim ctlAction, strProgramName, aElmIDList(80)
    Dim strNAChkBoxes, intNATop, aCtlIDs(70), strPrgIDList
    Dim dctUnSorted, dtmEndDate, dtmStartDate
        
    If strWhoCalled <> "BuildTab" Then
        Form.TabBuildCompleted.value = "N^"
    End If

    Set dctBuild = CreateObject("Scripting.Dictionary")
    Set dctPrgs = CreateObject("Scripting.Dictionary")
    Set dctUnSorted = CreateObject("Scripting.Dictionary")
    strPrograms = "^"
    intJ = 0
    For intI = 1 To 5
        document.all("lblFunction" & intI).innerText = ""
        document.all("lblStatus" & intI).style.left = -1000
        document.all("lblFunction" & intI).style.left = -1000
    Next

    For Each oDictObj In mdctPrograms
        blnProgramSelected = False
        intProgramID = oDictObj
        If document.all("chkProgram" & intProgramID).checked = True Then
            blnProgramSelected = True
            If Not IsReviewFull(document.all("cboReviewType" & intProgramID).value) Then
                strRecord = mdctReviewTypes(CLng(document.all("cboReviewType" & intProgramID).value))
                aElmIDList(intProgramID) = Parse(strRecord,"^",2)
            Else
                aElmIDList(intProgramID) = "All"
            End If
        End If
        If blnProgramSelected = True Then
            strPrograms = strPrograms & intProgramID & "^"
            dctPrgs.Add CLng(intProgramID), 580 + (120*intJ) + (10*intJ)
            intJ = intJ + 1
            document.all("lblFunction" & intJ).innerText = mdctPrograms(CLng(intProgramID))
            document.all("lblFunction" & intJ).style.textAlign = "center"
            document.all("lblFunction" & intJ).style.fontWeight = "bold"
            intK = intJ - 1
            document.all("lblFunction" & intJ).style.left = 580 + (intK*120) + (intK*10)
            document.all("lblStatus" & intJ).style.left = 580 + (intK*120) + (intK*10)

            If intJ = 5 Then Exit For
        End If
    Next
    
    For Each oElm In window.opener.mdctElements
        strRecord = window.opener.mdctElements(oElm)
        blnIncludeFull = False
        If UCase(Parse(strRecord,"^",6)) = "TRUE" Then
            blnIncludeFull = True
        End If
        
        dtmEndDate = Parse(strRecord,"^",3)
        dtmStartDate = Parse(strRecord,"^",7)
        If dtmEndDate = "" Then dtmEndDate = "12/31/2100"
        If InStr(strPrograms,"^" & Parse(strRecord,"^",4) & "^") > 0 And CInt(Parse(strRecord,"^",5)) = 2 And CDate(txtReviewDateEntered.value) <= CDate(dtmEndDate) And CDate(txtReviewDateEntered.value) >= CDate(dtmStartDate) Then            
            If InStr(aElmIDList(Parse(strRecord,"^",4)),"[" & oElm & "]") > 0 Or (aElmIDList(Parse(strRecord,"^",4)) = "All" And blnIncludeFull) Then
                strScreenName = Parse(strRecord,"^",1)
                strFactorList = RemoveInactiveFactors(Parse(strRecord,"^",8))
                If strFactorList <> "0.*" Then
                    For intJ = 1 To 100
                        strFactor = Parse(strFactorList,"*",intJ)
                        If strFactor = "" Then Exit For
                        intFactorID = Parse(strFactor,"*",1)
                        If Not dctUnSorted.Exists(strScreenName & "^" & intFactorID) Then
                            dctUnSorted.Add strScreenName & "^" & intFactorID, Parse(strRecord,"^",4) & "^"
                        Else
                            dctUnSorted(strScreenName & "^" & intFactorID) = dctUnSorted(strScreenName & "^" & intFactorID) & Parse(strRecord,"^",4) & "^"
                        End If
                    Next
                End If
            End If
        End If
    Next
    
    strHoldScreen = ""
    For intI = 1 To 1000    
        For Each oDictObj In dctUnSorted
            strHoldScreen = Parse(oDictObj,"^",1)
            Exit For
        Next
        For Each oDictObj In dctUnSorted
            strScreenName = Parse(oDictObj,"^",1)
            If strHoldScreen = strScreenName Then
                dctBuild.Add oDictObj, dctUnSorted(oDictObj)
                dctUnSorted.Remove oDictObj
            End If
        Next
        If dctUnSorted.Count = 0 Then Exit For
    Next
    intI = 0
    intK = 0
    intTop = 5
    strBuild = ""
    strHoldScreen = ""

    For Each oDictObj In dctBuild
        strScreenName = Parse(oDictObj,"^",1)
        If strHoldScreen <> strScreenName Then
            If strHoldScreen <> "" Then
                strPrgIDList = ""
                For intJ = 2 To 11
                    intProgramID = Parse(strNAChkBoxes,"|",intJ)
                    If intProgramID = "" Then Exit For
                    strPrgIDList = strPrgIDList & Parse(intProgramID,"^",1) & "F"
                    strBuild = strBuild & AddScreenNAControl(Parse(intProgramID,"^",1), intI, Parse(intProgramID,"^",2), Parse(intProgramID,"^",3), aCtlIDs(CInt(Parse(intProgramID,"^",1))))
                Next
                strRecord = AddScreenCommentHTML(intI, strHoldScreen, intTop, strPrgIDList)
                intTop = Parse(strRecord,"^",1)
                strBuild = strBuild & Parse(strRecord,"^",2)
            End If
            intI = intI + 1
            strRecord = AddScreenHTML(intI, strScreenName, intTop)
            intNATop = intTop - 1
            intTop = Parse(strRecord,"^",1)
            strBuild = strBuild & Parse(strRecord,"^",2)
            strHoldScreen = strScreenName
            strNAChkBoxes = "|"
            For intJ = 1 To 70
                aCtlIDs(intJ) = ""
            Next
        End If
        strProgramList = dctBuild(oDictObj)
        strBuild = strBuild & AddFieldNameHTML(0, intI, Parse(oDictObj,"^",2), intTop, 0, True, 0, strScreenName)
        For intJ = 1 To 10
            intProgramID = Parse(strProgramList,"^",intJ)
            If intProgramID = "" Then Exit For
            intK = intK + 1
            If InStr(strNAChkBoxes,"|" & intProgramID & "^") = 0 Then
                strNAChkBoxes = strNAChkBoxes & intProgramID & "^" & intNATop & "^" & dctPrgs(CLng(intProgramID)) + 20 & "|"
            End If
            aCtlIDs(intProgramID) = aCtlIDs(intProgramID) & intK & "^"
            strBuild = strBuild & AddFieldNameHTML(intProgramID, intI, Parse(oDictObj,"^",2), intTop, dctPrgs(CLng(intProgramID)), False, intK, strScreenName)
        Next
        intTop = intTop + 20
    Next
    If strHoldScreen <> "" Then
        strPrgIDList = ""
        For intJ = 2 To 10
            intProgramID = Parse(strNAChkBoxes,"|",intJ)
            If intProgramID = "" Then Exit For
            strPrgIDList = strPrgIDList & Parse(intProgramID,"^",1) & "F"
            strBuild = strBuild & AddScreenNAControl(Parse(intProgramID,"^",1), intI, Parse(intProgramID,"^",2), Parse(intProgramID,"^",3), aCtlIDs(CInt(Parse(intProgramID,"^",1))))
        Next
        strRecord = AddScreenCommentHTML(intI, strHoldScreen, intTop, strPrgIDList)
        intTop = Parse(strRecord,"^",1)
        strBuild = strBuild & Parse(strRecord,"^",2)
    End If
    strBuild = strBuild & " <INPUT id=txtDataIntegrityInfo Type=""hidden"" Value=""" & intI & "^" & intK & """>"
    divDataIntegrity.innerHTML = strBuild
    
    If strWhoCalled <> "BuildTab" Then
        Call DisableTabControls(chkProgram1.disabled)
        Form.TabBuildCompleted.value = "Z^" & strWhoCalled
    End If
End Sub

Function AddScreenNAControl(intPrgID, intRowID, intTop, intLeft, strCtlIDList)
    Dim strBuild
    
    strBuild = ""
    strBuild = strBuild & "<INPUT type=""button"" ID=cmdScreenNAPrg" & intPrgID & "R" & intRowID & " onclick=""Call cmdScreenNA_onClick(" & intPrgID & "," & intRowID & ")"" style=""LEFT:" & intLeft & ";width:100;TOP:" & intTop - 3 & """ NAME=cmdScreenNAPrg" & intPrgID & "R" & intRowID & " value=""All Factors NA"">"
    strBuild = strBuild & "<INPUT type=""hidden"" ID=txtScreenNAPrg" & intPrgID & "R" & intRowID & " NAME=txtScreenNAPrg" & intPrgID & "R" & intRowID & " VALUE=""" & strCtlIDList & """>"
    
    AddScreenNAControl = strBuild
End Function

Function AddFieldNameHTML(intPrgID, intRowID, intFactorID, intTop, intLeft, blnShowFieldName, intCtlID, strScreenName)
    Dim intI
    Dim strBuild, strChecked
    Dim strRecord, strFacList
    Dim intElmID, strFacStatus

    If blnShowFieldName = True Then
        strBuild = "<SPAN id=lblElement" & intRowID & "CF" & intFactorID & " class=DefLabel style=""LEFT:15; WIDTH:568; TOP:" & intTop & ";OVERFLOW:HIDDEN;"""
        strBuild = strBuild & " onclick=""Call FactorOnClick(" & intFactorID & ",2)"" "
        strBuild = strBuild & "<B>" & GetFactorTitle(intFactorID) & "</B></SPAN>"
    Else
        intElmID = window.opener.mdctElementIDs(intPrgID & "^" & strScreenName)
        strFacStatus = "25"
        If mdctElmData.Exists(intPrgID & "^2^" & intElmID) Then
            strFacList = Parse(mdctElmData(intPrgID & "^2^" & intElmID),"*",4)
            For intI = 1 To 100
                strRecord = Parse(strFacList,"!",intI)
                If strRecord = "" Then Exit For
                If CLng(Parse(strRecord,"~",1)) = CLng(intFactorID) Then
                    strFacStatus = Parse(strRecord,"~",2)
                    If strFacStatus = "0" Then strFacStatus = "25"
                    Exit For
                End If
            Next
        End If

        For intI = 0 To 3
            strChecked = ""
            If intI + 22 = CInt(strFacStatus) Then strChecked = "checked"
            strBuild = strBuild & "<INPUT type=radio name=""optDataIntC" & intI & "R" & intCtlID & """ disabled onclick=""Call DataIntegrityField_onclick(" & intI & "," & intCtlID & ")""" 
            strBuild = strBuild & "style=""background-color:<%=gstrBackColor%>;LEFT:" & intLeft + (intI*25) + 21 & "; TOP:" & intTop & """ " & strChecked & " ID=optDataIntC" & intI & "R" & intCtlID & ">"
        Next
        strBuild = strBuild & "<INPUT type=""hidden"" id=txtDataIntR" & intCtlID & " value=""" & intPrgID & "^" & intRowID & "^" & intFactorID & """>"
    End If
    AddFieldNameHTML = strBuild
End Function

Function AddScreenCommentHTML(intRowID, strScreenName, intTop, strPrgIDList)
    Dim strBuild
    
    strBuild = strBuild & "<TEXTAREA id=txtCommentsType2Row" & intRowID & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:15; WIDTH:680; TOP:" & intTop & "; HEIGHT:30; TEXT-ALIGN: left; padding-left:3"""
    strBuild = strBuild & " onblur=""Call ElementCommentDI_onblur(" & intRowID & ")"" tabIndex=<%=GetTabIndex%> NAME=""txtCommentsType2Row" & intRowID & """>"
    If mdctElmComments.Exists(strScreenName) Then
        strBuild = strBuild & CleanTextRecordParsers(mdctElmComments(strScreenName),"FromDb","All")
    End If
    strBuild = strBuild & "</TEXTAREA>"
    strBuild = strBuild & "<INPUT id=txtDIProgramIDList" & intRowID & " Type=""hidden"" Value=""" & strPrgIDList & """>"
    AddScreenCommentHTML = intTop + 35 & "^" & strBuild
End Function

Function AddScreenHTML(intRowID, strScreenName, intTop)
    Dim strBuild
    
    strBuild = "<SPAN id=lblElement" & intRowID & " class=DefLabel style=""LEFT:10; WIDTH:400; TOP:" & intTop & "; HEIGHT:20; OVERFLOW:visible;"">"
    strBuild = strBuild & "<B>" & strScreenName & "</B></SPAN>"
    
    AddScreenHTML = intTop + 15 & "^" & strBuild
End Function

Function AddProgramHTML(intProgramID, intTop, intElementTypeID)
    Dim strBuild

    strBuild = "<SPAN id=lblTab" & intElementTypeID & "Program" & intProgramID & " class=DefLabel style=""LEFT:10; WIDTH:300; TOP:" & intTop & "; HEIGHT:20; OVERFLOW:visible; COLOR:<%=gstrTitleColor%>;FONT-WEIGHT:bold; FONT-SIZE:12pt;"">"
    strBuild = strBuild & "<B>" & mdctPrograms(intProgramID) & "</B>"
    strBuild = strBuild & "</SPAN>"
    strBuild = strBuild & "<SPAN id=lblTabEmpty" & intElementTypeID & "Program" & intProgramID & " class=DefLabel style=""LEFT:310; WIDTH:200; TOP:" & intTop & "; HEIGHT:20; OVERFLOW:visible; COLOR:<%=gstrTitleColor%>;FONT-WEIGHT:bold; FONT-SIZE:12pt;""></SPAN>"
    AddProgramHTML = intTop + 20 & "^" & strBuild
End Function

Function AddElementHTML(intProgramID, intTop, intElementTypeID)
    Dim intI, intRowTop, intRowID, intLeft, intDivHeight
    Dim strBuild, strBuildInfo
    Dim oDictObj, oTab, strFacKey
    Dim strRecord, strElmRecord

    intI = 0
    intRowTop = 0
    intRowID = 0
    strBuild = ""
    For Each oTab In mdctElmData
        If intProgramID & "^" & intElementTypeID = Parse(oTab,"^",1) & "^" & Parse(oTab,"^",2) Then
            mintDivsTabIndex = mintDivsTabIndex + 25
            strRecord = mdctElmData(oTab)
            strElmRecord = Parse(oTab,"^",3) & "*" & mdctElmData(oTab)
            intRowID = intRowID + 1
            strBuildInfo = ElementHTLMBuild(intProgramID, intRowTop, intElementTypeID, strElmRecord, intRowID, 0)
            strBuild = strBuild & Parse(strBuildInfo,"^",2)
            intRowTop = intRowTop + 50 + Parse(strBuildInfo,"^",1)
        End If
    Next
    intDivHeight = intRowTop + 100
    'Add a blank DIV
    For intI = 1 To 10
        If intI = 1 And intRowID < maElementOptions(intProgramID,intElementTypeID, 1) Then
            intLeft = 0
        Else
            intLeft = -1000
        End If
        intRowID = intRowID + 1
        mintDivsTabIndex = mintDivsTabIndex + 25
        strBuildInfo = ElementHTLMBuild(intProgramID, intRowTop, intElementTypeID, "0*0*0**0~", intRowID, intLeft)
        strBuild = strBuild & Parse(strBuildInfo,"^",2)
        If intI = 1 Then 
            If intLeft = 0 Then
                intRowTop = intRowTop + 50 + Parse(strBuildInfo,"^",1)
            Else
                intRowTop = intRowTop + 50 + Parse(strBuildInfo,"^",1)
            End If
        End If
    Next
    If intDivHeight < 220 And intElementTypeID = 1 Then intDivHeight = 220
    
    strBuild = "<DIV id=divTab" & intElementTypeID & "Prg" & intProgramID & " class=DefRectangle style=""border-style:none;LEFT:0; WIDTH:725; TOP:" & intTop & ";height:" & intDivHeight & ";OVERFLOW:auto;background-color:transparent"">" & strBuild
    strBuild = strBuild & "<INPUT id=txtProgramInfoType" & intElementTypeID & "Program" & intProgramID & " style=""top:0;width:10:left:0"" Type=""hidden"" Value=""" & intRowID & """>"
    strBuild = strBuild & "</DIV>"

    AddElementHTML = intTop + intRowTop & "^" & strBuild
End Function

Function ElementHTLMBuild(intProgramID, intRowTop, intElementTypeID, strElementRecord, intRowID, intDivLeft)
    Dim strBuild, strRecord
    Dim intWidth, strOptions, intLeft, intTop
    Dim intElementID, strValue, strKey
    Dim strDecDiv, intDecIndex

    intElementID = Parse(strElementRecord, "*", 1)
    strRecord = window.opener.mdctElements(CLng(intElementID))
    
    If intElementTypeID = 1 Then
        intWidth = 295
        strOptions = "<%=maOptions(1)%></SELECT>" 
    Else
        intWidth = 100
        strOptions = "<option value=0></option><option value=22>Yes</option><option value=23>No</option></SELECT>" '"<%=maOptions(3)%></SELECT>"
    End If
    
    strKey = "Type" & intElementTypeID & "Prg" & intProgramID & "Row" & intRowID
    strBuild = ""
    strBuild = strBuild & "<SELECT id=cboAction" & strKey & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:10; TOP:0; WIDTH:400; OVERFLOW:auto"" TabIndex=" & mintDivsTabIndex
    strBuild = strBuild & " onchange=""Call Action_onchange(" & intElementTypeID & "," & intRowID & "," & intProgramID & ")"" NAME=""cboActionType" & intElementTypeID & "Row" & intRowID & """>"
    strBuild = strBuild & Replace(maElementOptions(intProgramID,intElementTypeID, 0),"<option value=" & intElementID & ">","<option selected value=" & intElementID & ">") & "</SELECT>"

    'Add a DIV for Decisions, regardless if current element has them or not.  This will be used
    'to add/clear decisions when the Action dropdown changes.
    strRecord = AddDecisionHTML(RemoveInactiveFactors(Parse(strRecord,"^",8)), Parse(strElementRecord,"*",5), intElementID, intElementTypeID, intRowID, intProgramID)
    strDecDiv = "<DIV id=divFactors" & strKey & " class=DefRectangle style=""border-style:none;LEFT:30; WIDTH:605; TOP:23;height:" & Parse(strRecord,"^",1) & ";OVERFLOW:visible;background-color:transparent"">" & Parse(strRecord,"^",2) & "</DIV>"
    strBuild = strBuild & strDecDiv
    strBuild = strBuild & " <INPUT id=txtElementInfoType" & intElementTypeID & "Prg" & intProgramID & "Row" & intRowID & " style=""top:0;width:10:left:0"" Type=""hidden"" Value=""" & intElementID & """>"
    intTop = 23 + CInt(Parse(strRecord,"^",1))

    mintDivsTabIndex = mintDivsTabIndex + 15
    strBuild = strBuild & "<SELECT id=cboStatus" & strKey & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:10; TOP:" & intTop & "; WIDTH:" & intWidth & "; OVERFLOW:auto"" TabIndex=" & mintDivsTabIndex
    strBuild = strBuild & " onchange=""Call ElementStatus_onchange(" & intElementTypeID & "," & intRowID & "," & intProgramID & ")"" NAME=""cboStatusType" & intElementTypeID & "Row" & intRowID & """>"
    strValue = Parse(strElementRecord,"*",2)
    If CInt(intElementID) = CInt(mintArrearageID) Then
        strOptions = "<%=maOptions(4)%></SELECT>"
        strBuild = strBuild & Replace(strOptions, "<option value=" & strValue & ">","<option selected value=" & strValue & ">")
    Else
        strBuild = strBuild & Replace(strOptions, "<option value=" & strValue & ">","<option selected value=" & strValue & ">")
    End If
    If intElementTypeID = 1 Then
        mintDivsTabIndex = mintDivsTabIndex + 1
        strBuild = strBuild & "<SELECT id=cboTimeFrame" & strKey & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:310; TOP:" & intTop & "; WIDTH:150; OVERFLOW:auto"" TabIndex=" & mintDivsTabIndex
        strBuild = strBuild & " disabled onchange=""Call ElementTimeframe_onchange(" & intElementTypeID & "," & intRowID & "," & intProgramID & ")"" NAME=""cboTimeFrameType" & intElementTypeID & "Row" & intRowID & """>"
        strValue = Parse(strElementRecord,"*",3)
        strBuild = strBuild & Replace("<%=maOptions(2)%>", "<option value=" & strValue & ">","<option selected value=" & strValue & ">")
        strBuild = strBuild & "</SELECT>" 
        intWidth = 260
        intLeft = 465
    Else
        intWidth = 508
        intLeft = 120
    End If
    mintDivsTabIndex = mintDivsTabIndex + 1
    strValue = CleanTextRecordParsers(Parse(strElementRecord,"*",4),"FromDb","All")
    If intElementTypeID = 1 Then
        strBuild = strBuild & "<TEXTAREA id=txtComments" & strKey & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:" & intLeft & "; WIDTH:" & intWidth & "; TOP:0; HEIGHT:" & intTop+20 & "; TEXT-ALIGN: left; padding-left:3"""
    Else
        strBuild = strBuild & "<TEXTAREA id=txtComments" & strKey & " disabled style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:" & intLeft & "; WIDTH:" & intWidth & "; TOP:23; HEIGHT:20; TEXT-ALIGN: left; padding-left:3"""
    End If
    strBuild = strBuild & " onblur=""Call ElementComment_onblur(" & intElementTypeID & "," & intRowID & "," & intProgramID & ")"" tabIndex=" & mintDivsTabIndex & " NAME=""txtCommentsType" & intElementTypeID & "Row" & intRowID & """>" & strValue & "</TEXTAREA>"
    strBuild = strBuild & "</DIV>"
    
    If intDivLeft = 0 Then
        strBuild = "<DIV id=div" & strKey & " class=DefRectangle style=""border-style:none;LEFT:0; WIDTH:705; TOP:" & intRowTop & ";height:" & intTop+30 & ";OVERFLOW:visible;background-color:transparent"">" & strBuild
    Else
        strBuild = "<DIV id=div" & strKey & " class=DefRectangle style=""border-style:none;LEFT:-1000; WIDTH:705; TOP:0;height:56;OVERFLOW:visible;background-color:transparent"">" & strBuild
    End If
    
    ElementHTLMBuild = intTop & "^" & strBuild
End Function
   
Function AddDecisionHTML(strAvailFac, strSelFac, intElementID, intElementTypeID, intRowID, intProgramID)
    Dim intFactorHeight, strKey
    Dim strOptions, intI, intJ
    Dim intFactorID, strFRecord, blnOk, intFactorRespID
    Dim strBuild
    
    strKey = "Type" & intElementTypeID & "Prg" & intProgramID & "Row" & intRowID
    intFactorHeight = 0
    If strAvailFac <> "0*" And intElementID <> 0 Then
        strOptions = "<option value=0></option><option value=22>Yes</option><option value=23>No</option><option value=24>NA</option>"
        'Action has decisions associated with it
        For intI = 1 To 100
            intFactorID = Parse(strAvailFac,"*",intI)
            If intFactorID = "" Then Exit For
            'Check if factor has a value selected
            intFactorRespID = 0
            For intJ =  1 To 100
                strFRecord = Parse(strSelFac,"!",intJ)
                If strFRecord = "" Then Exit For
                If CInt(Parse(strFRecord,"~",1)) = CInt(intFactorID) Then
                    intFactorRespID = Parse(strFRecord,"~",2)
                    Exit For
                End If
            Next
            strBuild = strBuild & "<SPAN id=lblFactor" & strKey & "F" & intFactorID & " title=""Click on Decision for a description"" class=DefLabel style=""cursor:hand;LEFT:0; WIDTH:310; TOP:" & intFactorHeight & ";OVERFLOW:hidden;"""
            strBuild = strBuild & " onmouseover=""Call FactorMouse(0," & intElementTypeID & "," & intRowID & "," & intProgramID & "," & intFactorID & ")"" onmouseout=""Call FactorMouse(1," & intElementTypeID & "," & intRowID & "," & intProgramID & "," & intFactorID & ")"" onclick=""Call FactorOnClick(" & intFactorID & "," & intElementTypeID & ")"">"
            strBuild = strBuild & "" & GetFactorTitle(intFactorID) & ""
            strBuild = strBuild & "</SPAN>"
            mintDivsTabIndex = mintDivsTabIndex + 1
            strBuild = strBuild & "<SELECT id=cboStatus" & strKey & "F" & intFactorID & " style=""BACKGROUND-COLOR:<%=gstrBackColor%>;LEFT:310;TOP:" & intFactorHeight & ";WIDTH:70; OVERFLOW:auto"" TabIndex=" & mintDivsTabIndex
            strBuild = strBuild & " onchange=""Call FactorStatus_onchange(" & intElementTypeID & "," & intRowID & "," & intProgramID & "," & intFactorID & ")"" NAME=""cboStatus" & strKey & "F" & intFactorID & """>"
            strBuild = strBuild & Replace(strOptions, "<option value=" & intFactorRespID & ">","<option selected value=" & intFactorRespID & ">")
            strBuild = strBuild & "</SELECT>"
            intFactorHeight = intFactorHeight + 20
        Next
        intFactorHeight = intFactorHeight
    End If
    
    AddDecisionHTML = intFactorHeight & "^" & strBuild
End Function

Function RemoveInactiveFactors(strFactorList)
    Dim intI
    Dim strNewList, strRecord
    
    strNewList = ""
    For intI = 1 To 100
        strRecord = Parse(strFactorList,"*",intI)
        If strRecord = "" Then Exit For
        If Parse(strRecord,".",2) <> "" Then
            If CDate(txtReviewDateEntered.value) <= CDate(Parse(strRecord,".",2)) Then
                strNewList = strNewList & Parse(strRecord,".",1) & "*"
            End If
        Else
            strNewList = strNewList & Parse(strRecord,".",1) & "*"
        End If
    Next
    RemoveInactiveFactors = strNewList
End Function

Sub AdjustActionAndProgramDivHeights(intTabID)
    Dim intI, intJ
    Dim intTop, strKey
    
    intTop = 10
    strKey = ""

    For intJ = 1 To mintPrgCount
        If document.all("chkProgram" & intJ).checked = True Then
            '   Check to ensure there is at least one action for the function.  If there isn't, skip
            If maElementOptions(intJ, intTabID, 1) > 0 Then
                strKey = "Type" & intTabID & "Prg" & intJ & "Row"
                document.all("lblTab" & intTabID & "Program" & intJ).style.Top = intTop
                document.all("lblTabEmpty" & intTabID & "Program" & intJ).style.Top = intTop
                document.all("divTab" & intTabID & "Prg" & intJ).style.posTop = intTop + 20
                For intI = 1 To document.all("txtProgramInfoType" & intTabID & "Program" & intJ).value
                    If document.all("div" & strKey & intI).style.posleft >= 0 Then
                        'do nothing
                    Else
                        'When first non-visible action div is found:
                        '   Use the top+height of last visible action to determine the height of the next program div.
                        intTop = intTop + document.all("div" & strKey & intI-1).style.posTop + document.all("div" & strKey & intI-1).style.posHeight + 40
                        Exit For
                    End If
                Next
                document.all("lblTabEmpty" & intTabID & "Program" & intJ).innerText = ""
            Else
                document.all("lblTabEmpty" & intTabID & "Program" & intJ).innerText = "**No Questions**"
            End If
        End If
    Next
End Sub

Sub Action_onchange(intTabID, intRowID, intProgramID)
    Dim strKey, strmdctElmDataKey
    Dim intResp, intI, intTop
    Dim intElementID, strAvailFactors
    Dim strRecord
    Dim strKeyStub
    
    For intI = 1 To document.all("txtProgramInfoType" & intTabID & "Program" & intProgramID).value
        strKey = "Type" & intTabID & "Prg" & intProgramID & "Row"
        If CLng(document.all("cboAction" & strKey & intI).value) > 0 And _
            CInt(intI) <> CInt(intRowID) And _
            CLng(document.all("cboAction" & strKey & intI).value) = CLng(document.all("cboAction" & strKey & intRowID).value) Then
            'Trying to change the Action to the same action selected somewhere else.  Action can only be selected 1 time per program
            If intTabID = 1 Then
                MsgBox "This Action has been previously selected.",vbOkOnly,"Action Change"
            Else
                MsgBox "This Question has been previously selected.",vbOkOnly,"Information Gathering"
            End If
            'document.all("cboAction" & strKey & intRowID).value = 0
            document.all("cboAction" & strKey & intRowID).value = document.all("txtElementInfo" & strKey & intRowID).value
            Exit Sub
        End If
    Next
    strKey = "Type" & intTabID & "Prg" & intProgramID & "Row" & intRowID
    If CLng(document.all("cboAction" & strKey).value) <> CLng(document.all("txtElementInfo" & strKey).value) Then
        If CLng(document.all("txtElementInfo" & strKey).value) <> 0 Then
            intResp = MsgBox("Changing the Action will clear the status, timeframe and any decisions.  Are you sure?",vbYesNo,"Action Change")
            If intResp = vbNo Then
                document.all("cboAction" & strKey).value = document.all("txtElementInfo" & strKey).value
                Exit Sub
            End If
        End If
    End If
    
    'If selected Action has decisions, add them
    intElementID = document.all("cboAction" & strKey).value
    mintDivsTabIndex = document.all("cboAction" & strKey).tabIndex

    strAvailFactors = RemoveInactiveFactors(Parse(window.opener.mdctElements(CLng(intElementID)),"^",8))
    strRecord = AddDecisionHTML(strAvailFactors,"",intElementID,intTabID, intRowID, intProgramID)
    document.all("divFactors" & strKey).innerHTML = Parse(strRecord,"^",2)
    document.all("divFactors" & strKey).style.height = Parse(strRecord,"^",1)
    'Increase height of Action div to account for newly added decisions
    document.all("div" & strKey).style.Height = 45 + Parse(strRecord,"^",1)
    document.all("cboStatus" & strKey).style.top = 23 + Parse(strRecord,"^",1)
    If intTabID = 1 Then
        document.all("cboTimeFrame" & strKey).style.top = 23 + Parse(strRecord,"^",1)
        document.all("txtComments" & strKey).style.top = 0
        document.all("txtComments" & strKey).style.height = 43 + Parse(strRecord,"^",1)
    End If
    
    strKeyStub = "Type" & intTabID & "Prg" & intProgramID & "Row"
    If CInt(intRowID) < CInt(document.all("txtProgramInfoType1Program" & intProgramID).value) Then
        If document.all("div" & strKeyStub & intRowID+1).style.posLeft < 0 Then
            'If next action div is not visible, make it visible unless the max number of actions
            'has been selected.
            If maElementOptions(intProgramID,intTabID, 1) > CInt(intRowID) Then
                document.all("div" & strKeyStub & intRowID+1).style.left = 0
            End If
        End If
        'Adjust the top of any visible div's below current one and programs div height/top
        'intHeightChange = CInt(document.all("div" & strKey).style.posHeight) - CInt(intOrigHeight)
        For intI = intRowID+1 To document.all("txtProgramInfoType" & intTabID & "Program" & intProgramID).value
            If document.all("div" & "Type" & intTabID & "Prg" & intProgramID & "Row" & intI).style.posleft >= 0 Then
                intTop = document.all("div" & strKeyStub & intI-1).style.posTop + document.all("div" & strKeyStub & intI-1).style.posHeight + 10
                document.all("div" & strKeyStub & intI).style.top = intTop
            Else
                'When first non-visible action div is found:
                '   Use the top+height of last visible action to determine the height of the program div.
                document.all("divTab" & intTabID & "Prg" & intProgramID).style.height = document.all("div" & strKeyStub & intI-1).style.posTop + document.all("div" & strKeyStub & intI-1).style.posHeight + 40
                Exit For
            End If
        Next
    Else
        ' If last available div is being displayed, increase program div here
        document.all("divTab" & intTabID & "Prg" & intProgramID).style.height = document.all("div" & strKeyStub & intRowID).style.posTop + document.all("div" & strKeyStub & intRowID).style.posHeight + 40
    End If
    
    ' If current row is second to last available div, the above IF statement will not increase the program div,
    ' so do it here.
    If CInt(intRowID+1) = CInt(document.all("txtProgramInfoType1Program" & intProgramID).value) Then
        document.all("divTab" & intTabID & "Prg" & intProgramID).style.height = document.all("div" & strKeyStub & intRowID+1).style.posTop + document.all("div" & strKeyStub & intRowID+1).style.posHeight + 40
    End If
    intTop = document.all("divTab" & intTabID & "Prg" & intProgramID).style.posTop + document.all("divTab" & intTabID & "Prg" & intProgramID).style.posheight + 20
    If intProgramID < mintPrgCount Then
        For intI = intProgramID + 1 To mintPrgCount
            If document.all("chkProgram" & intI).checked = True Then
                document.all("lblTab" & intTabID & "Program" & intI).style.Top = intTop
                document.all("lblTabEmpty" & intTabID & "Program" & intI).style.Top = intTop
                document.all("divTab" & intTabID & "Prg" & intI).style.Top = intTop + 20
                intTop = document.all("divTab" & intTabID & "Prg" & intI).style.posTop + document.all("divTab" & intTabID & "Prg" & intI).style.posheight + 20
            End If
        Next
    End If
    
    strmdctElmDataKey = intProgramID & "^" & intTabID
    If CLng(document.all("cboAction" & strKey).value) = 0 Then
        Call mdctElmData_Remove(strmdctElmDataKey & "^" & document.all("txtElementInfo" & strKey).value) 
        If CInt(document.all("txtElementInfo" & strKey).value) = CInt(mintArrearageID) Then
            Call ReFillActionStatusDropDown(False,strKey)
        End If
        If CInt(intProgramID) = 6 Then
            'If action that is being removed is under enfor rem, the action name has a "function" equivalent and
            'we must clear any data integrity items from the dictionary object.  They will be keyed under the "function"
            'equivalent ID.
            Call RemoveFunctionFromDictionary("E^" & document.all("txtElementInfo" & strKey).value)
        End If
        document.all("txtElementInfo" & strKey).value = 0
        document.all("cboStatus" & strKey).value = 0
        If intTabID = 1 Then
            document.all("cboTimeFrame" & strKey).value = 0
            document.all("cboTimeFrame" & strKey).disabled = True
        End If
        document.all("txtComments" & strKey).value = ""
    Else
        If CLng(document.all("txtElementInfo" & strKey).value) <> 0 Then
            Call mdctElmData_Remove(strmdctElmDataKey & "^" & document.all("txtElementInfo" & strKey).value) 
        End If
        Call mdctElmData_Add(strmdctElmDataKey & "^" & document.all("cboAction" & strKey).value, strAvailFactors) 
        document.all("txtElementInfo" & strKey).value = document.all("cboAction" & strKey).value
        If CInt(document.all("cboAction" & strKey).value) = CInt(mintArrearageID) Then
            Call ReFillActionStatusDropDown(True,strKey)
        Else
            Call ReFillActionStatusDropDown(False,strKey)
        End If
    End If
    If CInt(intProgramID) = 6 And CInt(intTabID) = 1 Then
        document.all("cboAction" & strKey).disabled = True
        Form.TabBuildCompleted.value = "A^" & strKey
        Form.TabsDisabled.value = "True"
        mlngTimerIDB = window.setInterval("CheckForBuildCompletion",100)
    End If
End Sub

Sub ReFillActionStatusDropDown(blnArrearage, strKey)
    Dim oOption, ctlDropDown
    Dim intI
    Dim strRecord, strRecords
    
    If Mid(strKey,5,1) <> "1" Then Exit Sub
    
    Set ctlDropDown = document.all("cboStatus" & strKey)
    ctlDropDown.options.length = Null
    Set oOption = Document.createElement("OPTION")
	    oOption.Value = 0
	    oOption.Text = ""
	    ctlDropDown.options.Add oOption
    Set oOption = Nothing
     
    If blnArrearage Then
        strRecords = "<%=maOptions(6)%>"
    Else
        strRecords = "<%=maOptions(5)%>"
    End If

    For intI = 1 To 10
        strRecord = Parse(strRecords,"|",intI)
        If strRecord = "" Then Exit For
	    Set oOption = Document.createElement("OPTION")
		    oOption.Value = Parse(strRecord,"^",1)
		    oOption.Text = Parse(strRecord,"^",2)
		    ctlDropDown.options.Add oOption
	    Set oOption = Nothing
    Next
End Sub

Sub mdctElmData_Add(strKey, strFactorList)
    Dim intI

    <%'Check to see if Element exists in Item list for the Key. (It should not, but just to be sure.)  If it does, 
    'delete it.  Add new element. %>
    If mdctElmData.Exists(strKey) Then
        mdctElmData.Remove(strKey)
    End If
    mdctElmData.Add strKey, "0*0**" & Replace(strFactorList,"*","~0!")
End Sub

Sub mdctElmData_UpdateElement(strKey, intStatusID, intTimeFrameID, strComments)
    Dim intI
    Dim strRecord

    <%'Check to see if Element exists in Item list for the Key.%>
    If mdctElmData.Exists(strKey) Then
        strRecord = mdctElmData(strKey)
        ' Only element data updated here, so re-add factor list as is
        strRecord = intStatusID & "*" & intTimeFrameID & "*" & CleanTextRecordParsers(strComments,"ToDb","All") & "*" & Parse(strRecord,"*",4)
        
        mdctElmData(strKey) = strRecord
    Else
        <%'If no record exists for key (program^tab), add a new item to dict obj with a blank factor list. (Should never get here) %>
        mdctElmData.Add strKey, intStatusID & "*" & intTimeFrameID & "*" & CleanTextRecordParsers(strComments,"ToDb","All") & "*0~"
    End If
End Sub

Sub mdctElmData_UpdateFactor(strKey, intFactorID, intFactorStatusID)
    Dim intI
    Dim strElmRecord, strRecord
    Dim strFacRecord, strNewFacRecord

    <%'Check to see if Element exists.  If it does,  
    'rebuild the item with all existing element data EXCEPT the factor list. 
    'Add updated factor info to the end of the current factor list. %>
    If mdctElmData.Exists(strKey) Then
        strElmRecord = mdctElmData(strKey)
        ' Only a single factor record updated here, so re-add element info and all other factors as is
        strRecord = Parse(mdctElmData(strKey),"*",4)
        For intI = 1 To 100
            strFacRecord = Parse(strRecord,"!",intI)
            If strFacRecord = "" Then Exit For
            If CInt(Parse(strFacRecord,"~",1)) <> CInt(intFactorID) Then
                strNewFacRecord = strNewFacRecord & strFacRecord & "!"
            End If
        Next
        strNewFacRecord = strNewFacRecord & intFactorID & "~" & intFactorStatusID & "!"
        
        mdctElmData(strKey) = Parse(strElmRecord,"*",1) & "*" & Parse(strElmRecord,"*",2) & "*" & Parse(strElmRecord,"*",3) & "*" & strNewFacRecord
    Else
        <%'If no record exists for key (program^tab), add a new item to dict obj with a single factor in list. (Should never get here) %>
        mdctElmData.Add strKey, "0*0**" & intFactorID & "~" & intFactorStatusID & "!"
    End If
End Sub

Sub mdctElmData_Remove(strKey)
    <%'Check to see if Element exists.  If it does, and it should, remove it.%>
    If mdctElmData.Exists(strKey) Then
        mdctElmData.Remove(strKey)
    End If
End Sub

Sub mdctElmData_Display(strKey)
    Dim intI, oElm, intJ
    Dim strElmRecord, strRecord, strItem, strFactorList, strFactor

    strItem = ""
    For Each oElm In mdctElmData
        If oElm = strKey Or strKey = "All" Then
            strItem = strItem & "KEY:" & oElm & vbCrLf
            
            strRecord = mdctElmData(oElm)
            strItem = strItem & Space(5) & "EID:" & Parse(oElm,"^",3) & ".."
            strItem = strItem & "EStID:" & Parse(strRecord,"*",1) & ".."
            strItem = strItem & "ETfID:" & Parse(strRecord,"*",2) & ".."
            strItem = strItem & "EComm:" & Parse(strRecord,"*",3) & ".." & vbCrLf
            strFactorList = Parse(strRecord,"*",4)
            For intJ = 1 To 20
                strFactor = Parse(strFactorList,"!",intJ)
                If strFactor = "" Then Exit For
                If Parse(strFactor,"~",1) <> "0" Then
                    strItem = strItem & Space(10) & "FID(" & intJ & "):" & Parse(strFactor,"~",1) & "..StID:" & Parse(strFactor,"~",2) & vbCrLf
                End If
            Next
            strItem = strItem & "-------------------------------------" & vbCrLf
        End If
    Next
    MsgBox strItem
End Sub

Sub cmdScreenNA_onClick(intPrgID, intRowID)
    Dim intI
    Dim strCtlID

    For intI = 1 To 100
        strCtlID = Parse(document.all("txtScreenNAPrg" & intPrgID & "R" & intRowID).value,"^",intI)
        If strCtlID = "" Then Exit For
        document.all("optDataIntC2R" & strCtlID).checked = True
        Call DataIntegrityField_onclick(2,strCtlID)
    Next
End Sub

Sub lblProgram_onClick(intRowID)
    If document.all("chkProgram" & intRowID).disabled = True Then Exit Sub
    document.all("chkProgram" & intRowID).checked = Not document.all("chkProgram" & intRowID).checked
    Call chkProgram_onclick(intRowID)
End Sub

Sub FactorOnClick(intFactorID, intTabID)
    Select Case intTabID
        Case 1
        Case 2
            lblDIFactorDescr.innerText = GetFactorDescription(intFactorID)
    End Select
End Sub

Sub FactorMouse(intDir, intTabID, intRowID, intPrgID, intFactorID)
    If intDir = 0 Then
        document.all("lblFactorType" & intTabID & "Prg" & intPrgID & "Row" & intRowID & "F" & intFactorID).style.fontweight = "bold"
    Else
        document.all("lblFactorType" & intTabID & "Prg" & intPrgID & "Row" & intRowID & "F" & intFactorID).style.fontweight = "normal"
    End If
End Sub

<%'priority sorts either by "Value" Or "Text", depending on which is passed %>
Sub SortcboList(cboList,strPriority)
    Dim strTemp1
    Dim strTemp2
    Dim arrList(2,25)
    
    Dim intI
    Dim intJ
    
    For intI = 0 to 25 Step 1
        arrList(0, intI) = 0
        arrList(1, intI) = ""
        arrList(2, intI) = 0
    Next
    
    For intI = 0 To cboList.options.Length - 1 Step 1
        arrList(0 , intI) = Parse(cboList.Options.Item(intI).Value, ":", 1)        
        arrList(1 , intI) = Parse(cboList.Options.Item(intI).Text, ":", 1)
        arrList(2 , intI) = Parse(cboList.Options.Item(intI).Value, ":", 2)
    Next
    
    If strPriority  = "Value" Then
        For intI = 0 To cboList.options.Length Step 1
            For intJ = 0 To cboList.options.Length Step 1
                If cint(arrList(0, intJ)) > cint(arrList(0, intJ + 1)) Then
                    If Not cint(arrList(0, intJ)) = 0 Or cint(arrList(0, intJ + 1)) = 0 Then
                        strTemp1 = cint(arrList(0, intJ + 1))
                        strTemp2 = arrList(1, intJ + 1)
                        
                        arrList(0, intJ + 1) = cint(arrList(0, intJ))
                        arrList(1, intJ + 1) = arrList(1, intJ)
                        
                        arrList(0, intJ) = cint(strTemp1)
                        arrList(1, intJ) = strTemp2
                    End If
                End If
           Next
        Next
        
        intJ = 0
        For intI = 0 To 25 Step 1
            If Not arrList(0, intI) = 0 Then
                cboList.Options.Item(intJ).Value = arrList(0 , intI) & ":" & arrList(2 , intI)
                cboList.Options.Item(intJ).Text = arrList(1 , intI)
                intJ = intJ + 1
            End If
            
        Next

    ElseIf strPriority  = "Text" Then
        <%'Step through list beginning to end%>
        For intI = 0 To cboList.options.Length Step 1
            For intJ = 0 To cboList.options.Length Step 1
                If strComp(LCase(arrList(1, intJ)), LCase(arrList(1, intJ + 1)), 1) = 1 Then
                    strTemp1 = arrList(0, intJ + 1)
                    strTemp2 = arrList(1, intJ + 1)
                    
                    arrList(0, intJ + 1) = arrList(0, intJ)
                    arrList(1, intJ + 1) = arrList(1, intJ)
                    
                    arrList(0, intJ) = strTemp1
                    arrList(1, intJ) = strTemp2
                End If
            Next
        Next
        intJ = 0
        For intI = 0 To 25 Step 1
            If Not arrList(0, intI) = 0 Then
                cboList.Options.Item(intJ).Value = arrList(0 , intI)& ":" & arrList(2 , intI)
                cboList.Options.Item(intJ).Text = arrList(1 , intI)
                intJ = intJ + 1
            End If
        Next
    End If
End Sub

<%'----------------------------------------------------------------------------
' Name:    MoveListOption()
' Purpose: This subroutine moves a combo Or listbox (SELECT) option from the
'          source list to the destination list.  
'----------------------------------------------------------------------------%>
Sub MoveListOption(ctlSource, intWhich, ctlDest)
    Dim oOption
    
	Set oOption = Document.createElement("OPTION")
		oOption.Value = ctlSource.options.Item(intWhich).Value
		oOption.Text = ctlSource.options.Item(intWhich).Text
		ctlDest.options.Add oOption
	Set oOption = Nothing
    ctlSource.options.Remove(intWhich)
End Sub

Sub CheckCaseStatus()
    Dim oElm
    Dim strStatus, strElmData
    
    strStatus = "Correct"
    For Each oElm In mdctElmData
        strElmData = mdctElmData(oElm)
        If Parse(oElm,"^",2) = "1" And Left(strElmData,2) <> "30" Then
            strStatus = "<%=gstrErrorTitle%>"
            Exit For
        End If
        If Parse(oElm,"^",2) = "2" And InStr(strElmData,"~23!") > 0 Then
            strStatus = "<%=gstrErrorTitle%>"
            Exit For
        End If
        If Parse(oElm,"^",2) = "3" And Left(strElmData,2) <> "22" Then
            strStatus = "<%=gstrErrorTitle%>"
            Exit For
        End If
    Next
    txtCaseStatus.value = strStatus
End Sub

Sub cmdGuide_onclick()
    If GuideWindowIsSet Then
        oGuideWindow.focus
        Exit Sub
    End If
    
    Call OpenGuideWindow
End Sub

Sub ClearGuideWindow()
    Set oGuideWindow = Nothing
End Sub

Sub OpenGuideWindow()
    Dim intX
    Dim intY
    Dim dteStart

    intX = window.screenLeft + 487
    intY = window.screenTop
End Sub

Sub LoadAuditDictionary()
    Dim strURL
    strURL = "ActivityAudit.asp?Action=Read&RecordID=" & Form.rvwID.value & "&Table=tblReviews"

    Set mdctAudit = CreateObject("Scripting.Dictionary")
    Set mdctAudit = window.showModalDialog(strURL)
End Sub

Sub DisplayAuditActivity()
    Dim oAudit
    Dim strRecord
    Dim strInnerHTML, strOuterHTML
    Dim intI
    Dim strCursor, strTitle
    
    intI = InStr(tblAudit.outerHTML,"<TBODY")
    If intI > 0 Then
        strOuterHTML = Left(tblAudit.outerHTML,intI-1)
        tblAudit.outerHTML = strOuterHTML & " <TBODY id=tbdAudit></TBODY></TABLE>"
    End If

    strOuterHTML = tblAudit.outerHTML
    
    For Each oAudit In mdctAudit
        strRecord = mdctAudit(oAudit)
        If InStr(Parse(strRecord,"^",4),"Update") > 0 Then
            strCursor = "hand"
            strTitle = "Double-Click for details"
        Else
            strCursor = "default"
            strTitle = ""
        End If
        strInnerHTML = strInnerHTML & "<TR id=tdrAudit" & oAudit & " title=""" & strTitle & """ style=""cursor:" & strCursor & """>" & vbCrLf
        strInnerHTML = strInnerHTML & "    <TD class=TableDetail onclick=SelectRow(" & oAudit & ") ondblclick=DisplayAuditDetails(" & oAudit & ") id=tdcAuditC0" & oAudit & " style=""width:100;text-align:center"">" & Parse(strRecord,"^",3) & "</TD>" & vbCrLf
        strInnerHTML = strInnerHTML & "    <TD class=TableDetail onclick=SelectRow(" & oAudit & ") ondblclick=DisplayAuditDetails(" & oAudit & ") id=tdcAuditC1" & oAudit & " style=""width:140;text-align:center"">" & Parse(strRecord,"^",2) & "</TD>" & vbCrLf
        If Parse(strRecord,"^",4) <> "Update" Then
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail onclick=SelectRow(" & oAudit & ") ondblclick=DisplayAuditDetails(" & oAudit & ") id=tdcAuditC2" & oAudit & " style=""width:300;padding-left:5"">" & Parse(strRecord,"^",5) & "</TD>" & vbCrLf
        Else
            strInnerHTML = strInnerHTML & "    <TD class=TableDetail onclick=SelectRow(" & oAudit & ") ondblclick=DisplayAuditDetails(" & oAudit & ") id=tdcAuditC2" & oAudit & " style=""width:300;padding-left:5"">Review Updated</TD>" & vbCrLf
        End If
        strInnerHTML = strInnerHTML & "</TR>" & vbCrLf
    Next
    tblAudit.outerHTML = Replace(strOuterHTML,"<TBODY id=tbdAudit></TBODY>","<TBODY id=tbdAudit>" & strInnerHTML & "</TBODY>")
End Sub

Sub SelectRow(intID)
    Dim objRow
    
    For Each objRow In tblAudit.rows
        If objRow.ID = "tdrAudit" & intID Then
            objRow.className = "TableSelectedRow"
        Else
            objRow.className = "TableRow"
        End If
    Next
End Sub

Sub DisplayAuditDetails(intID)
    Dim oAudit
    Dim strRecord
    Dim strReturnValue

    For Each oAudit In mdctAudit
        If CLng(oAudit) = CLng(intID) Then
            strRecord = mdctAudit(oAudit)
            If InStr(Parse(strRecord,"^",4),"Update") > 0 Then
                strReturnValue = window.showModalDialog("ActivityAudit.asp?Action=Details" & _
                    "&Details=" & Replace(Parse(strRecord,"^",5),"#","[PDSGN]") & _
                    "&ChangeDate=" & Parse(strRecord,"^",2) & _
                    "&RecordID=" & Parse(strRecord,"^",7) & _
                    "&UserID=" & Parse(strRecord,"^",3) _
                    ,, "dialogWidth:610px;dialogHeight:420px;scrollbars:no;center:yes;border:thin;help:no;status:no")
            End If
            Exit For
        End If
    Next
End Sub

<%'----------------------------------------------------------------------------
' Name:    GuideWindowIsSet()
' Purpose: This function is called when the guide button is clicked to check
'          whether or not the guide window is already open.
'----------------------------------------------------------------------------%>
Function GuideWindowIsSet()
End Function

</SCRIPT>

<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=PageBody>

    <DIV id=divFormTitle CLASS=ReviewTitleArea
        STYLE="HEIGHT:32; WIDTH:434; PADDING-TOP:4; BORDER-STYLE:solid; BORDER-WIDTH:2; BORDER-RIGHT-WIDTH:1">
        <SPAN id=lblFormTitle class=DefLabel style="LEFT:90; WIDTH:330; TOP:4;font-size:13;text-align:right"><%= mstrPageTitle & " ~ Enter Case Review"%></SPAN>
        <DIV id=divNavigateButton style="left:4;top:4;font-size:13;width:75">
            Navigate
        </DIV>
    </DIV>

    <% Call WriteNavigateControls(1,0,Null) %>
    
    <DIV id=divReviewID CLASS=ReviewTitleArea
        STYLE="TOP: 1; LEFT:434; HEIGHT:32; WIDTH:320; BORDER-STYLE:solid; BORDER-WIDTH:2; BORDER-LEFT-WIDTH:1">

        <SPAN id=lblReviewID class=DefLabel style="LEFT:10; WIDTH:65; TOP:4">
            Review ID
            <INPUT id=txtCaseReviewID type=text title="Case Review Record ID"
                style="LEFT:65; WIDTH:65; BACKGROUND-COLOR:<%=gstrAltBackColor%>"
                tabIndex=<%=GetTabIndex%> readOnly NAME="txtCaseReviewID">
        </SPAN>
        <SPAN id=lblReviewEnteredDate class=DefLabel style="LEFT:155; WIDTH:80; TOP:4">
            Review Date
            <INPUT id=txtReviewDateEntered type=text title="Review Date Entered"
                style="LEFT:80; WIDTH:70; background-color:<%=gstrAltBackColor%>"
                tabIndex=<%=GetTabIndex%> readOnly NAME="txtReviewDateEntered">
        </SPAN>
    </DIV>
        
    <DIV ID=SaveWindow Name=SaveWindow 
        style="position:absolute;
            TOP:29; 
            LEFT:-1000;
            WIDTH:754; 
            HEIGHT:453; 
            COLOR:black; 
            BORDER-STYLE:solid;
            BORDER-WIDTH:2;
            BORDER-COLOR:<%=gstrBorderColor%>;
            BACKGROUND-COLOR:<%=gstrBackColor%>">
        <IFRAME ID=SaveFrame src="Blank.html" Name=SaveFrame style="top:0;height:400;left:0;width:750">
        </IFRAME>
        <SPAN id=lblSavingMessage class=DefLabel style="top:110;height:45;left:0;width:750"><CENTER><BIG><B></BIG>Saving Case Review Record...</B></BIG></CENTER></SPAN>
    </DIV> 

    <!-- IFRAME used for searching for staff----------------------------->
    <DIV id=divStaffSearch class=ControlDiv
        style="width:200;height:150;left:-1000;top:40;z-index:101">
        <IFRAME id=fraStaffSearch src="Blank.html" tabindex=-1 style="width:198;height:148;left:1;top:1"></IFRAME>
    </DIV>
    <DIV id=divCaseBody CLASS=DefPageFrame STYLE="LEFT:1; TOP:36; WIDTH:753; HEIGHT:453; visibility:hidden">
        <%'-- Header portion of the case review --%>
        <DIV id=divReviewHeader STYLE="TOP:0; WIDTH:754; HEIGHT:140; OVERFLOW:hidden">

            <%'-- First row of controls ------------------------------------------%>
            <DIV id=divRow1 STYLE="TOP:5; WIDTH:754; HEIGHT:18; OVERFLOW:visible; z-index:-1">

                <SPAN id=lblReviewer class=DefLabel style="LEFT:8; WIDTH:75">
                    <%=gstrRvwTitle%>
                    <INPUT type="text" ID=txtReviewer NAME="txtReviewer" readonly 
                        style="LEFT:70; WIDTH:200; TOP:1;BACKGROUND-COLOR:<%=gstrBackColor%>">
                </SPAN>
                <SPAN id=lblReviewMonthYear class=DefLabel style="LEFT:315; WIDTH:80">
                    Review Month
                    <INPUT type=text id=txtReviewMonthYear title="Review Month/Year" style="LEFT:80; WIDTH:90"
                        tabIndex=<%=GetTabIndex%> disabled maxlength=7 NAME="txtReviewMonthYear">
                </SPAN>
                
                <SPAN id=lblReviewClass class=DefLabel style="LEFT:520; Width:75">
                    <% = gstrReviewClassTitle %>
                    <SELECT id=cboReviewClass title="<% = gstrReviewClassTitle %>"
                        style="LEFT:75; WIDTH:145; Height:50"  
                        tabIndex=<%=GetTabIndex%> disabled=True NAME="cboReviewClass">
                        <OPTION value=0></OPTION>
                        <%=BuildList("ReviewClass",Null,0,0,0)%>
                    </SELECT>
                </SPAN>
            </DIV><%'-- End First row of controls --%>

            <%'-- Second row of controls ------------------------------------------%>
            <DIV id=divRow2 STYLE="TOP:32; WIDTH:754; HEIGHT:18; OVERFLOW:visible; z-index:-1">

                <SPAN id=lblCaseNumber class=DefLabel style="LEFT:8; WIDTH:125">
                    Case / Referral Number
                </SPAN>
                <INPUT type=text id=txtCaseNumber title="Case Number" style="LEFT:130; WIDTH:200"
                    onfocus="CmnTxt_onfocus(txtCaseNumber)"
                    tabIndex=<%=GetTabIndex%> disabled maxlength=<%=mintMaxCaseNumLen%> NAME="txtCaseNumber">

                <SPAN id=lblClientName class=DefLabel style="LEFT:375; WIDTH:75">
                    Case Name
                </SPAN>
                <INPUT type=text id=txtClientLastName title="Case Client Last Name" style="LEFT:445; WIDTH:120"
                    onfocus="CmnTxt_onfocus(txtClientLastName)"
                    tabIndex=<%=GetTabIndex%> maxlength=50 disabled NAME="txtClientLastName">
                <INPUT id=txtClientFirstName title="Case Client First Name" style="LEFT:570; WIDTH:75"
                    onfocus="CmnTxt_onfocus(txtClientFirstName)"
                    tabIndex=<%=GetTabIndex%> maxlength=50 disabled NAME="txtClientFirstName">
            </DIV><%'-- End Second row of controls --%>

            <%'-- Third row of controls ------------------------------------------%>
            <DIV id=divRow3 STYLE="TOP:48; WIDTH:754; HEIGHT:18; OVERFLOW:visible; z-index:-1">
            
                <SPAN id=lblWorker class=DefLabel style="LEFT:8; WIDTH:175;TOP:5">
                    <%=gstrWkrTitle%> ID / Name  
                    <INPUT type="text" title="Worker Caseload ID" ID=txtWorkerID NAME="txtWorkerID" 
                        onkeydown="Gen_onkeydown(txtWorkerID)"
                        onfocus="CmnTxt_onfocus(txtWorkerID)" tabIndex=-1 maxlength=20
                        style="LEFT:-1000; WIDTH:80; TOP:14;BACKGROUND-COLOR:<%=gstrBackColor%>">
                     <INPUT type="text" title="Worker ID" ID=txtWorkerEmpID NAME="txtWorkerEmpID" 
                        onkeydown="Gen_onkeydown(txtWorkerEmpID)"
                        onfocus="CmnTxt_onfocus(txtWorkerEmpID)" tabIndex=<%=GetTabIndex%> maxlength=20
                        style="LEFT:0; WIDTH:60; TOP:14;BACKGROUND-COLOR:<%=gstrBackColor%>">
                     <INPUT type="text" title="Worker Name" ID=txtWorker NAME="txtWorker" 
                        onfocus="CmnTxt_onfocus(txtWorker)" 
                        tabIndex=<%=GetTabIndex%> style="LEFT:62; WIDTH:125; TOP:14;BACKGROUND-COLOR:<%=gstrBackColor%>">
                </SPAN>

                <SPAN id=lblSupervisor class=DefLabel style="LEFT:200; WIDTH:170;TOP:5">
                    <%=gstrSupTitle%> ID / Name
                </SPAN>
                <INPUT type="text" title="Supervisor Worker ID" ID=txtSupervisorEmpID NAME="txtSupervisorEmpID" 
                    onkeydown="Gen_onkeydown(txtSupervisorEmpID)"
                    onfocus=CmnTxt_onfocus(txtSupervisorEmpID)
                    tabIndex=<%=GetTabIndex%> maxlength=20
                    style="LEFT:200; WIDTH:60; TOP:19;BACKGROUND-COLOR:<%=gstrBackColor%>">
                <INPUT type="text" title="Supervisor Name" id=txtSupervisor
                    style="LEFT:265; WIDTH:125; TOP:19;BACKGROUND-COLOR:<%=gstrBackColor%>"
                        onfocus="CmnTxt_onfocus(txtSupervisor)" 
                    tabIndex=<%=GetTabIndex%>  NAME="txtSupervisor">
                    
                <SPAN id=lblManager class=DefLabel style="LEFT:-1395; WIDTH:170;TOP:5">
                    <%=gstrMgrTitle%>
                </SPAN>
                <SELECT id=cboManager style="LEFT:-1395; WIDTH:180;top:19;BACKGROUND-COLOR:<%=gstrBackColor%>" tabIndex=<%=GetTabIndex%> NAME="cboManager">
                    <option value=""></option>
                </SELECT>

                <SPAN id=lblOffice class=DefLabel style="LEFT:-1590; WIDTH:70;TOP:5">
                    <%=gstrOffTitle%>
                </SPAN>
                <SELECT id=cboOffice style="LEFT:-1590; WIDTH:150;top:19;BACKGROUND-COLOR:<%=gstrBackColor%>" tabIndex=<%=GetTabIndex%> NAME="cboOffice">
                    <option value=""></option>
                </SELECT>
            </DIV><%'-- End Third row of controls --%>
        </DIV><%'-- Header portion of the case review --%>

        <%'-----Eligibility Element Display Area ---------------------------------%>
        <DIV id=divTabButton1 class="defRectangle DivTab" style="LEFT:0;top:90;width:250" onclick="divTabs_onclick(1)" onkeydown="divTabs_onkeydown(1)" tabIndex=<%=GetTabIndex%>>
            Programs
        </DIV>
        <DIV id=divTab1 class=defRectangle style="LEFT:0; TOP:108; WIDTH:751; HEIGHT:305; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>;overflow:auto">
            <SPAN id=lblFunctions class=DefLabel style="LEFT:10; WIDTH:280; TOP:5">
                <B>Select Programs to Review</B>
            </SPAN>
            <SPAN id=lblReviewTypes class=DefLabel style="LEFT:220; WIDTH:280; TOP:5">
                <B>Review Type</B>
            </SPAN>
            <DIV id=divFunctionsLoading class=DefRectangle style="LEFT:-1010; WIDTH:545;height:210;overflow:auto;TOP:25;border-style:none;BACKGROUND-COLOR:<%=gstrAltBackColor%>">
<%
                mintI = 0
                For Each moDictObj In mdctPrograms
                    Response.Write "<INPUT type=""checkbox"" ID=chkProgramD" & moDictObj & " style=""LEFT:1;TOP:" & mintI*22 & """ disabled NAME=chkProgramD" & moDictObj & ">"
                    Response.Write "<SPAN id=lblProgramD" & moDictObj & " class=DefLabel style=""color:gray;LEFT:21; WIDTH:180; TOP:" & mintI*22 & """>" & mdctPrograms(moDictObj) & "</SPAN>"
                    Response.Write "<SELECT id=cboReviewTypeD" & moDictObj & " disabled style=""LEFT:210; TOP:" & mintI*22 & "; WIDTH:180; OVERFLOW:auto"" TabIndex=" & GetTabIndex & vbCrLf
                    Response.Write "NAME=""cboReviewTypeD" & moDictObj & """>" & vbCrLf
                    Response.Write "<option value=""55"">Full</option></SELECT>"
                    mintI = mintI + 1
                Next
%>
            </DIV>
            <DIV id=divFunctions class=DefRectangle style="LEFT:10; WIDTH:545;height:210;overflow:auto;TOP:25;border-style:none;BACKGROUND-COLOR:<%=gstrAltBackColor%>">
<%
                mintI = 0
                For Each moDictObj In mdctPrograms
                    Response.Write "<INPUT type=""checkbox"" ID=chkProgram" & moDictObj & " onclick=chkProgram_onClick(" & moDictObj & ") style=""LEFT:1;TOP:" & mintI*22 & """ NAME=chkProgram" & moDictObj & ">"
                    Response.Write "<SPAN id=lblProgram" & moDictObj & " onclick=lblProgram_onClick(" & moDictObj & ") class=DefLabel style=""cursor:hand;LEFT:21; WIDTH:180; TOP:" & mintI*22 & """>" & mdctPrograms(moDictObj) & "</SPAN>"
                    Response.Write "<SELECT id=cboReviewType" & moDictObj & " onchange=cboReviewType_onClick(" & moDictObj & ") disabled style=""LEFT:210; TOP:" & mintI*22 & "; WIDTH:180; OVERFLOW:auto"" TabIndex=" & GetTabIndex & vbCrLf
                    Response.Write "NAME=""cboReviewType" & moDictObj & """>" & vbCrLf
                    Response.Write "<option value=""0""></option>"
                    Response.Write "<option value=""55"">Full</option></SELECT>"
                    mintI = mintI + 1
                Next
%>
            </DIV>
        </DIV><%'-----End Eligibility Element Display Area --%>

        <%'-----Program Specific Tab ----------------------------------------------%>
        <DIV id=divTabButton2 class="defRectangle DivTab" style="LEFT:250;top:90;width:252" onclick="divTabs_onclick(2)" onkeydown="divTabs_onkeydown(2)" tabIndex=<%=GetTabIndex%>>
            Elements
        </DIV>
        <DIV id=divTab2 class=defRectangle style="LEFT:-5000; TOP:108; WIDTH:751; HEIGHT:305; 
            OVERFLOW:auto; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>">
            <SPAN id=lblScreenName class=DefLabel style="LEFT:10; WIDTH:200; TOP:5; HEIGHT:20; OVERFLOW:visible">
            <B>Elements</B>
            </SPAN>
            <SPAN id=lblFunction1 class=DefLabel style="LEFT:260; WIDTH:135; TOP:5; HEIGHT:15; OVERFLOW:hidden;">
            <B></B>
            </SPAN>
            <SPAN id=lblFunction2 class=DefLabel style="LEFT:-405; WIDTH:135; TOP:5; HEIGHT:15; OVERFLOW:hidden;">
            <B></B>
            </SPAN>
            <SPAN id=lblFunction3 class=DefLabel style="LEFT:-550; WIDTH:135; TOP:5; HEIGHT:15; OVERFLOW:hidden;">
            <B></B>
            </SPAN>
            <SPAN id=lblFunction4 class=DefLabel style="LEFT:-695; WIDTH:135; TOP:5; HEIGHT:15; OVERFLOW:hidden;">
            <B></B>
            </SPAN>
            <SPAN id=lblFunction5 class=DefLabel style="LEFT:-830; WIDTH:135; TOP:5; HEIGHT:15; OVERFLOW:hidden;">
            <B></B>
            </SPAN>
            <SPAN id=lblStatus1 class=DefLabel style="LEFT:-405;text-align:center;WIDTH:135; TOP:20; HEIGHT:20; OVERFLOW:visible;">
            <B>Yes&nbsp;&nbsp;No&nbsp;&nbsp;&nbsp;NA&nbsp;&nbsp;&nbsp;NR</B>
            </SPAN>
            <SPAN id=lblStatus2 class=DefLabel style="LEFT:-550;text-align:center;WIDTH:135; TOP:20; HEIGHT:20; OVERFLOW:visible;">
            <B>Yes&nbsp;&nbsp;No&nbsp;&nbsp;&nbsp;NA&nbsp;&nbsp;&nbsp;NR</B>
            </SPAN>
            <SPAN id=lblStatus3 class=DefLabel style="LEFT:-550;text-align:center;WIDTH:135; TOP:20; HEIGHT:20; OVERFLOW:visible;">
            <B>Yes&nbsp;&nbsp;No&nbsp;&nbsp;&nbsp;NA&nbsp;&nbsp;&nbsp;NR</B>
            </SPAN>
            <SPAN id=lblStatus4 class=DefLabel style="LEFT:-550;text-align:center;WIDTH:135; TOP:20; HEIGHT:20; OVERFLOW:visible;">
            <B>Yes&nbsp;&nbsp;No&nbsp;&nbsp;&nbsp;NA&nbsp;&nbsp;&nbsp;NR</B>
            </SPAN>
            <SPAN id=lblStatus5 class=DefLabel style="LEFT:-550;text-align:center;WIDTH:135; TOP:20; HEIGHT:20; OVERFLOW:visible;">
            <B>Yes&nbsp;&nbsp;No&nbsp;&nbsp;&nbsp;NA&nbsp;&nbsp;&nbsp;NR</B>
            </SPAN>
<%
            Response.Write "<DIV id=divDataIntegrity class=DefRectangle style=""border-style:none;LEFT:0; WIDTH:750; TOP:40;height:225;OVERFLOW:auto;background-color:transparent"">" & vbCrLf
            Response.Write "</DIV>"
%>
            <SPAN id=lblDIFactorDescr class=DefLabel style="border-style:solid;border-width:1;LEFT:10; WIDTH:730; TOP:270; HEIGHT:30; OVERFLOW:auto;">
            </SPAN>
        </DIV><%'-----Program Specific Tab --%>

        <%'-----Case Summary Tab --------------------------------------------------%>
        <DIV id=divTabButton3 class="defRectangle DivTab" style="LEFT:502;top:90;width:250" onclick="divTabs_onclick(3)" onkeydown="divTabs_onkeydown(3)" tabindex=-1>
            Review Action
        </DIV>
        <DIV id=divTab3 class=defRectangle style="LEFT:175; TOP:108; WIDTH:751; HEIGHT:305; OVERFLOW:auto; BORDER-RIGHT-STYLE:none; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrAltBackColor%>">
            <SPAN id=lblRvwComments class=DefLabel style="LEFT:320; WIDTH:415; TOP:7">
                Supervisor Comments
            </SPAN>
            <BUTTON id=cmdExpandRvwComments onclick=ExpandReviewComments(cmdExpandRvwComments)
                STYLE="LEFT:711; WIDTH:20; TOP:2; HEIGHT:20; FONT-SIZE:6pt" 
                tabIndex=<%=GetTabIndex%>>
                /\
            </BUTTON>
            <TEXTAREA id=txtRvwComments style="LEFT:320; WIDTH:410; TOP:22; HEIGHT:45; TEXT-ALIGN:left; padding-left:3; overflow:auto"
                tabIndex=<%=GetTabIndex%> NAME="txtRvwComments"></TEXTAREA>

            <SPAN id=lblRvwCommentsWkr class=DefLabel style="LEFT:320; WIDTH:415; TOP:74">
                Worker Comments
            </SPAN>
            <BUTTON id=cmdExpandRvwCommentsWkr onclick=ExpandReviewComments(cmdExpandRvwCommentsWkr)
                STYLE="LEFT:711; WIDTH:20; TOP:68; HEIGHT:20; FONT-SIZE:6pt" 
                tabIndex=<%=GetTabIndex%>>
                /\
            </BUTTON>
            <TEXTAREA id=txtRvwCommentsWkr style="LEFT:320; WIDTH:410; TOP:90; HEIGHT:45; TEXT-ALIGN:left; padding-left:3; overflow:auto"
                tabIndex=<%=GetTabIndex%> NAME="txtRvwCommentsWkr"></TEXTAREA>
  
           <DIV id=divCheckBoxes class=defRectangle STYLE="TOP:5; WIDTH:310; HEIGHT:165;LEFT:5; OVERFLOW:visible;BACKGROUND-COLOR:transparent">
                <SPAN id=lblSignature class=DefLabel style="LEFT:0; WIDTH:133; TOP:5;text-align:center">
                    <B>Review Signatures</B>
                </SPAN>
                <SPAN id=lblSignature1 class=DefLabel style="cursor:hand;LEFT:2; WIDTH:125; TOP:30;TEXT-ALIGN:left" unselectable=on
                    onclick=SignatureOnClick(1)>Supervisor Signature</SPAN>
                <INPUT id=chkSignature1 style="LEFT:120; WIDTH:14; HEIGHT:14; TOP:29"
                    onclick=SignatureOnClickCtl(1) type=checkbox disabled tabIndex=<%=GetTabIndex%> NAME="chkSignature1">
                <SPAN id=lblResponse class=DefLabel style="LEFT:140; WIDTH:195;top:10;TEXT-ALIGN:left">
                    <%=gstrWkrTitle%> Response Requirement
                    <SELECT id=cboResponse title="<%=gstrWkrTitle%> response" style="LEFT:0; WIDTH:150;TOP:15"
                        tabIndex=<%=GetTabIndex%> disabled NAME="cboResponse">
                        <OPTION VALUE=0 SELECTED>
                        <%=WriteOptionList("WorkerResponse")%>
                    </SELECT>
                </SPAN>
                <SPAN id=lblCorrectionDue class=DefLabel style="LEFT:140; WIDTH:80;TOP:45; TEXT-ALIGN: left">
                    Response Due
                    <INPUT type=text id=txtCorrectionDue title="Response Or Correction Due Date" style="LEFT:0; WIDTH:80;TOP:15"
                        tabIndex=<%=GetTabIndex%> disabled NAME="txtCorrectionDue">
                </SPAN>
                <SPAN id=lblSignature2 class=DefLabel style="cursor:hand;LEFT:2; WIDTH:125; TOP:92;TEXT-ALIGN:left" unselectable=on
                    onclick=SignatureOnClick(2)>Worker Signature</SPAN>
                <INPUT id=chkSignature2 style="LEFT:120; WIDTH:14; HEIGHT:14; TOP:101"
                    onclick=SignatureOnClickCtl(2) type=checkbox disabled tabIndex=<%=GetTabIndex%> NAME="chkSignature2">
                <SPAN id=lblResponseW class=DefLabel style="LEFT:140; WIDTH:95;top:84;TEXT-ALIGN:left">
                    <%=gstrWkrTitle%> Response
                    <SELECT id=cboResponseW title="<%=gstrWkrTitle%> response" style="LEFT:0; WIDTH:150;TOP:15"
                        tabIndex=<%=GetTabIndex%> disabled NAME="cboResponseW">
                        <OPTION VALUE=0 SELECTED>
                        <%=WriteOptionList("WorkerResponseWorker")%>
                    </SELECT>
                </SPAN>
                <SPAN id=lblSignature3 class=DefLabel style="cursor:hand;LEFT:2; WIDTH:125; TOP:140;TEXT-ALIGN:left" unselectable=on
                    onclick=SignatureOnClick(3)>Submit To Reports</SPAN>
                <INPUT id=chkSignature3 style="LEFT:120; WIDTH:14; HEIGHT:14; TOP:139"
                    onclick=SignatureOnClickCtl(3) type=checkbox disabled tabIndex=<%=GetTabIndex%> NAME="chkSignature3">
          
            </DIV>
            <DIV id=divRow5 STYLE="TOP:138; WIDTH:744; HEIGHT:25; OVERFLOW:visible">
                <SPAN id=lblCaseStatus class=DefLabel style="LEFT:-1008; WIDTH:40">
                    Status
                    <TEXTAREA id=txtCaseStatus title="Case Status"
                        style="LEFT:40; WIDTH:65; background-color:<%=gstrBackColor%>"
                        tabIndex=<%=GetTabIndex%> readOnly NAME="txtCaseStatus"></TEXTAREA>
                </SPAN>

                    

                <BUTTON id=cmdSaveRecord class=DefButton title="Save New Record Or Changes to Current Record" 
                    STYLE="LEFT:520; WIDTH:65; TOP:14; HEIGHT:20" 
                    disabled accesskey=S tabIndex=288>
                    <u>S</u>ave
                </BUTTON>

                <BUTTON id=cmdPrint class=DefButton title="Save and Preview Current Record"
                    STYLE="LEFT:592; WIDTH:65; TOP:14; HEIGHT:20" 
                    disabled accesskey=P tabIndex=289>
                    <u>P</u>review
                </BUTTON>

            </DIV><%'-----Bottom row of controls --%>
            <DIV id=divRvwHistory Class=TableDivArea style="LEFT:8; TOP:175; WIDTH:730; HEIGHT:125; 
                OVERFLOW:auto; FONT-WEIGHT:normal;z-index:1200">
                <TABLE id=tblAudit Border=0 Rules=rows Width=700 CellSpacing=0
                    Style="position:absolute;overflow: hidden; TOP:0">
                    <THEAD id=tbhAudit style="height:17">
                        <TR id=thrAudit>
                            <TD class=CellLabel id=thcAuditC0 style="width:100;padding-left:0;padding-right:0">User ID</TD>
                            <TD class=CellLabel id=thcAuditC1 style="width:140;padding-left:0;padding-right:0">Date</TD>
                            <TD class=CellLabel id=thcAuditC2 style="width:460;padding-left:0;padding-right:0">Action Taken</TD>
                        </TR>
                    </THEAD>
                    <TBODY id=tbdAudit>
                    </TBODY>
                </TABLE>
            </DIV>
        </DIV> <!-- End divTab3 -->

        <%'---------------------------------------------------------------%>
        <DIV id=fraButtons
            STYLE="LEFT:0; WIDTH:752; TOP:416; HEIGHT:35; border-style:solid; border-width:2; border-color:<%=gstrBorderColor%>; BACKGROUND-COLOR:<%=gstrAltBackColor%>">

            <SPAN id=lblDatabaseStatus
                class=DefLabel
                STYLE="VISIBILITY:hidden; LEFT:5; WIDTH:200; TOP:10; TEXT-ALIGN:center">
                Accessing Database...
            </SPAN>

            <BUTTON id=cmdFindRecord class=DefButton title="Find Record" 
                STYLE="LEFT:10; WIDTH:65; TOP:6; HEIGHT:20" 
                accesskey=F tabIndex=284>
                <U>F</U>ind Case 
            </BUTTON>

            <BUTTON id=cmdGuide title="Onine Review Guide"
                STYLE="LEFT:-1100; WIDTH:85; TOP:6; HEIGHT:20" 
                accesskey=G tabIndex=<%=GetTabIndex%>>
                Review <U>G</U>uide
            </BUTTON>

            <BUTTON id=cmdAddRecord class=DefButton title="Add a New Case Review Record" 
                STYLE="LEFT:210; WIDTH:65; TOP:6; HEIGHT:20" 
                accesskey=A tabIndex=285>
                <U>A</U>dd
            </BUTTON>

            <BUTTON id=cmdChangeRecord class=DefButton title="Modify the Current Record" 
                STYLE="LEFT:280; WIDTH:65; TOP:6; HEIGHT:20" 
                disabled accesskey=C tabIndex=285>
                <U>E</U>dit
            </BUTTON>
            
            <BUTTON id=cmdDeleteRecord class=DefButton title="Delete the Current Record" 
                STYLE="LEFT:350; WIDTH:65; TOP:6; HEIGHT:20" 
                disabled accesskey=D tabIndex=286>
                <U>D</U>elete
            </BUTTON>

            <BUTTON id=cmdCancelEdit class=DefButton title="Cancel Add Or Change" 
                STYLE="LEFT:420; WIDTH:65; TOP:6; HEIGHT:20"
                disabled accesskey=L tabIndex=287>
                Cance<U>l</U>
            </BUTTON>

            <BUTTON id=cmdClose class=DefButton title="Close the Data Entry Form" 
                STYLE="LEFT:675; WIDTH:65; TOP:6; HEIGHT:20" 
                tabIndex=290>
                Close
            </BUTTON>
        </DIV>
    </DIV>

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

    WriteFormField "AliasID", glngAliasPosID
    WriteFormField "UserAdmin", gblnUserAdmin
    WriteFormField "UserQA", gblnUserQA
    WriteFormField "SaveCompleted", ""
    WriteFormField "DeleteFail", ""
    WriteFormField "StaffInformation", ""
    WriteFormField "BaseTitle", Trim(gstrOrgAbbr & " " & gstrAppName)
    WriteFormField "SupervisorName", ""
    WriteFormField "ManagerName", ""
    WriteFormField "TabBuildCompleted", ""
    WriteFormField "TabsDisabled", ""
    WriteFormField "SupSubWorkerDisagree", ""

    Response.Write Space(4) & "</FORM>"
    Set gadoCmd = Nothing
    Set madoRs = Nothing
    gadoCon.Close
    Set gadoCon = Nothing
    %>
</BODY>
</HTML>
<!--#include file="IncWriteOptionList.asp"-->
<!--#include file="IncBuildList.asp"-->
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
<!--#include file="IncFormDefBlank.asp"-->
<!--#include file="IncFormDefEdit.asp"-->
<!--#include file="IncInteliType.asp"-->
<!--#include file="IncNavigateControls.asp"-->
