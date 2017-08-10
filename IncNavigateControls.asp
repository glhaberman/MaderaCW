<META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncNavigateControls.asp                                          '
' Purpose: This include file contains server side code to build the         '
'          MDI navigate menus.                                              '
'                                                                           '
'==========================================================================='

Sub WriteNavigateControls(intPage,intMenuTop,strBackgroundColor)
    Dim intTop
    Dim intHeight
    Dim intI
    Dim aPages(5,1)
    Dim intOptions
    
    intOptions = 5
    
    If intMenuTop = 0 Then intMenuTop = 23
    If IsNull(strBackgroundColor) Then strBackgroundColor = gstrAltBackColor
    
    aPages(0,0) = "lblMain"
    aPages(0,1) = "Main Menu"
    aPages(1,0) = "lblEnterCaseReviews"
    aPages(1,1) = "Enter Case Reviews"
    aPages(2,0) = "lblFindCaseReview"
    aPages(2,1) = "Find Case Review"
    aPages(3,0) = "lblNavReports"
    aPages(3,1) = "Reports"
    aPages(4,0) = "lblReReviews"
    aPages(4,1) = "Re-Reviews"
    aPages(5,0) = "lblFindReReview"
    aPages(5,1) = "Find Re-Reviews"
    
    intHeight = (intOptions * 18) + 10
    If intPage <= -1 Then intHeight = intHeight + 18
    
    Response.Write "<DIV id=divNavigateMenu style=""left:-1000;top:" & intMenuTop & ";height:" & intHeight & ";width:135;font-size:10;z-index:102;background-color:" & strBackgroundColor & "; BORDER-STYLE:solid; BORDER-WIDTH:1"">"
    intTop = 5
    For intI = 0 To intOptions
        If intI <> intPage And MenuItemSecurity(intI) Then
            Response.Write "    <SPAN id=" & aPages(intI,0) & " class=DefLabel Title = """ & aPages(intI,1) & """"
            Response.Write "        onmouseover=MenuMouseOver(" & aPages(intI,0) & ") onmouseout=MenuMouseOut(" & aPages(intI,0) & ") "
            Response.Write "        onclick=MenuOnClick(" & aPages(intI,0) & ") "
            Response.Write "        style=""LEFT:2;WIDTH:128;TOP:" & intTop & ";cursor:hand;height:18;z-index:90"">"
            Response.Write "        " & aPages(intI,1)
            Response.Write "    </SPAN>"
            intTop = intTop + 18
        Else
            ' Create SPAN as hidden so the setInterval below works
            Response.Write "    <SPAN id=" & aPages(intI,0) & " class=DefLabel Title = """ & aPages(intI,1) & """"
            Response.Write "        style=""visibility:hidden"">"
            Response.Write "    </SPAN>"
        End If
    Next
    Response.Write "</DIV>"
    If intPage = 3 Or intPage = -2 Then
        Response.Write "<INPUT type=""hidden"" ID=txtNavigateFix NAME=""txtNavigateFix"" value=""Y"">"
    Else
        Response.Write "<INPUT type=""hidden"" ID=txtNavigateFix NAME=""txtNavigateFix"" value=""N"">"
    End If
End Sub

Function MenuItemSecurity(intI)
    Dim strFind
    
    Select Case intI
        Case 1, 2 ' Edit, Find
            strFind = "[1]"
        Case 3    ' Reports
            strFind = "[2]"
        Case 4, 5 ' Regular Re-reviews
            strFind = "[3]"
        Case Else ' Main Menu
            strFind = "["
    End Select
    
    MenuItemSecurity = False
    If InStr(gstrOptions,strFind) > 0 Then MenuItemSecurity = True
End Function
'==========================================================================='
' Purpose: This section contains client side functions for navigating       '
'          between MDI windows.                                             '
'                                                                           '
'==========================================================================='
%>

<SCRIPT id=MDINavigate language=vbscript>
<!--
Dim mintCloseMenu
Dim sngInitialMenuOpen
Dim aLabels(5)

Sub divNavigateButton_onmouseover()
    If divNavigateMenu.style.left <> "-1000px" Or divNavigateButton.style.color = "gray" Then Exit Sub

    divNavigateMenu.style.left = 30
    sngInitialMenuOpen = Timer()

    aLabels(0) = "lblMain"
    aLabels(1) = "lblEnterCaseReviews"
    aLabels(2) = "lblFindCaseReview"
    aLabels(3) = "lblNavReports"
    aLabels(4) = "lblReReviews"
    aLabels(5) = "lblFindReReview"
    mintCloseMenu = Window.setInterval("CheckMenuDiv",500)
    
    If txtNavigateFix.value = "Y" Then
        Call NavigateFix("Open")
    End If
End Sub

Function CheckMenuDiv()
    Dim intI
    
    <%' Check if any menu options are highlighted.  If any are, exit function%>
    For intI = 0 To 5
        If document.all(aLabels(intI)).style.color = "white" Then
            Exit Function
        End If
    Next
    <%'If no menu options are highlighted, close the menu DIV unless it is 
      'within 2 seconds of menu being initially opened.%>
    If CSng(Timer()) - CSng(sngInitialMenuOpen) < 2 Then Exit Function
    Call CloseMenuDiv()
End Function

Sub CloseMenuDiv()
    divNavigateMenu.style.left = -1000
    Window.clearInterval(mintCloseMenu)
    If txtNavigateFix.value = "Y" Then
        Call NavigateFix("Close")
    End If
End Sub

Sub MenuMouseOver(oMenuItem)
    sngInitialMenuOpen = CSng(sngInitialMenuOpen) - 2 <%'Once a menu option is highlighted, change sngInitialMenuOpen to ensure menu is closed.%>
    oMenuItem.style.fontWeight = "bold"
    oMenuItem.style.color = "white"
    oMenuItem.style.backgroundcolor = "gray"
End Sub
Sub MenuMouseOut(oMenuItem)
    oMenuItem.style.fontWeight = "normal"
    oMenuItem.style.color = "black"
    oMenuItem.style.backgroundcolor = "transparent"
End Sub
Sub MenuOnClick(oMenuItem)
    Select Case oMenuItem.ID
        Case "lblMain"
            window.opener.focus
        Case "lblFindCaseReview"
            window.opener.Form.CalledFrom.Value = "CaseAddEdit.asp"
            window.opener.Form.action = "FindCase.asp"
            Call window.opener.ManageWindows(2,"Open")
        Case "lblNavReports"
            window.opener.Form.CalledFrom.Value = "Main"
            window.opener.Form.action = "Reports.asp"
            Call window.opener.ManageWindows(3,"Open")
        Case "lblEnterCaseReviews"
            window.opener.Form.CalledFrom.Value = "Main"
            window.opener.Form.Action = "CaseAddEdit.asp"
            Call window.opener.ManageWindows(1,"Open")
        Case "lblReReviews"
            window.opener.Form.CalledFrom.Value = "Main"
            window.opener.Form.Action = "ReReviewAddEdit.asp"
            Call window.opener.ManageWindows(4,"Open")
        Case "lblFindReReview"
            window.opener.Form.CalledFrom.Value = "Main"
            window.opener.Form.Action = "FindReReview.asp"
            Call window.opener.ManageWindows(5,"Open")
    End Select
    Call CloseMenuDiv()
End Sub

Function CheckForMain()
    Dim blnClosed
    Dim strName
    
    blnClosed = False
    On Error Resume Next
    strName = window.opener.name
    If Err.number <> 0 Then
        blnClosed = True
    End If
    On Error GoTo 0
    
    If blnClosed = True Then
        window.clearInterval(mintCheckForMain)
        ' Disable the Navigate button and hide menu div
        divNavigateButton.style.color = "gray"
        Call CloseMenuDiv()
        ' MainClosed() is a Sub on the page currently open
        Call MainClosed()
    End If
End Function

Sub CloseWindow(strTitle, blnSetCloseClicked, intMessage)
    Dim strMessage
    If strTitle = "" Then strTitle = "Main Menu Closed"
    
    If blnSetCloseClicked = True Then
        mblnCloseClicked = True
    End If
    Select Case intMessage
        Case 1
            strMessage = ""
        Case 2
            strMessage = "The record has been saved. " & Space(10) & vbCrLf
        Case 3
            strMessage = "The record has been deleted. " & Space(10) & vbCrLf
    End Select
    strMessage = strMessage & _
        "The Main Menu window was previously closed." & Space(10) & vbCrLf & _
        "This window will now be closed also.  If you" & vbCrLf & _
        "wish to continue using the application," & vbCrLf & _
        "please log on again."

    MsgBox strMessage, vbInformation, strTitle
    window.close
End Sub
-->
</script>