<%
'------------------------------------------------------------------------------
'       Name: IncGlobals.asp
'    Purpose: This file includes any delcarations and assignments that are 
'             needed on every page in the application.
'------------------------------------------------------------------------------
Dim gadoCon         'ADO Connection object used in all pages.
Dim gstrTitle       'Site-specific title to use on all page headings.
Dim gstrAppName     'Site-specific application name displayed in browser title bar.
Dim gintElmCnt      'Number of eligibility elements for the site.
Dim gintPrgCnt      'Number of programs for the site.
Dim gstrBackColor   'Background color for frames.
Dim gstrAltBackColor 'Alternate Background color for frames.
Dim gstrForeColor   'Foreground (text) color.
Dim gstrBorderColor 'Color for frame/box borders.
Dim gstrAccentColor 'Color for shadows or highlights.
Dim gstrPageColor   'Page background color.
Dim gstrTitleColor  'Color for title or heading text.
Dim gstrTitleFont   'Font family to use for titles.
Dim gstrTitleFontSize 'Font size in points of title text.
Dim gstrTitleFontSmallSize 'Smaller font size to user for titles.
Dim gstrCtrlBackColor 'Color for textbox background.
Dim gstrDefButtonColor 'Button background color.
Dim gstrDefButtonText 'Button text color.
Dim gstrOrgName     'Full name of client organization.
Dim gstrOrgAbbr     'Abbreviation for client organization.
Dim gstrLocationName 'State or County name
Dim gstrLocationAbbr 'State or County abbreviation
Dim gstrTextFont    'Default font for screen text.
Dim gadoCmd         'ADO command object reused.
Dim gadoRs          'ADO recordset reused.
Dim gstrErrorTitle  'Title to use for error, i.e. Incorrect.
Dim gstrUserID      'Holds the logged in user ID.
Dim gstrPassword    'Holds the users password
Dim glngUserStaffID 'Holds the employee ID of logged on user.
Dim glngPositionID  'Holds the user's current position ID
Dim glngAliasPosID  'Holds the Alias current position ID
Dim gstrUserName    'Name of the current logged on user
Dim gblnUserAdmin   'Administrative user flag.
Dim gblnUserQA      'Quality Assurance user flag.
Dim gblnUserRpts    'Reports user flag.
Dim gblnDebugOn     'Toggles on the inclusion of debugging functions.
Dim gstrOptions  	'Holds the security options assigned to the logged on user.
Dim gstrRoles		'Holds the Roles assigned to the logged on user.
Dim gstrWkrTitle	'Holds the title given to the worker management level
Dim gstrSupTitle	'Holds the title given to the Supervisor management level
Dim gstrMgrTitle	'Holds the title given to the Manager management level
Dim gstrOffTitle	'Holds the title given to the Office management level
Dim gstrDirTitle	'Holds the title given to the Director management level
Dim gstrRvwTitle	'Holds the title given to the Reviewer management level
Dim gstrAuthTitle	'Holds the title given to the Authorized By management level
Dim gstrEvaTitle	'Holds the title given to the Evaluator management level
Dim gstrEvaluation	'Holds the lable for an evaluation 
'Customer/version specific:
Dim gblnUserCAS     'CAS Reviewer user flag.
Dim gstrShrtTitle   'Holds the short title for the program
Dim gstrLngTitle    'Holds the long title for the program
Dim gstrWorkerRespDueDays    'Holds the number of days in the future for worker response
Dim gstrAllowReviewEdit 'Holds the Review edit flag
Dim gstrReviewClassTitle 'Holds title for Review Class
Dim gstrUseFactorCodes   
Dim gstrUserNamePosID
Dim gblnUseLogon    'Flag to switch between application logon and integrated security.

Function GetGlobalSettings()
    Dim adCmd
    Dim adRs
    Dim strSettingName
    Dim strSettingValue
    
    'Setting debug on will cause debugging functions to be included
    'in both the server-side and client-side common functions:
    gblnDebugOn = True   
    
    Set adCmd = GetAdoCmd("spGetAllSettings")
    Set adRs = GetAdoRs(adCmd)
    Do While Not adRs.EOF And Not adRs.BOF
        'Set any application wide settings and set global variables:
        strSettingName = adRs.Fields("SettingName").value
        strSettingValue = adRs.Fields("SettingValue").value

        Select Case strSettingName
            Case "MainFormTitle"
                gstrTitle = strSettingValue
            Case "ApplicationName"
                gstrAppName = strSettingValue
            Case "BackColor"
                gstrBackColor = strSettingValue
            Case "AltBackColor"
                gstrAltBackColor = strSettingValue
            Case "ForeColor"
                gstrForeColor = strSettingValue
            Case "BorderColor"
                gstrBorderColor = strSettingValue
            Case "AccentColor"
                gstrAccentColor = strSettingValue
            Case "PageColor"
                gstrPageColor = strSettingValue
            Case "TitleColor"
                gstrTitleColor = strSettingValue
            Case "TitleFont"
                gstrTitleFont = strSettingValue
            Case "TitleFontSize"
                gstrTitleFontSize = strSettingValue
            Case "TitleFontSmallSize"
                gstrTitleFontSmallSize = strSettingValue
            Case "CtrlBackColor"
                gstrCtrlBackColor = strSettingValue
            Case "DefButtonColor"
                gstrDefButtonColor = strSettingValue
            Case "DefButtonTextColor"
                gstrDefButtonText = strSettingValue
            Case "OrganizationName"
                gstrOrgName = strSettingValue
            Case "OrganizationAbbr"
                gstrOrgAbbr = strSettingValue
            Case "LocationName"
                gstrLocationName = strSettingValue
            Case "LocationAbbr"
                gstrLocationAbbr = strSettingValue
            Case "TextFont"
                gstrTextFont = strSettingValue
            Case "ErrorTitle"
                gstrErrorTitle = strSettingValue
            Case "ElementCount"
                gintElmCnt = strSettingValue
            Case "ProgramCount"
                gintPrgCnt = strSettingValue
            Case "WorkerTitle" 
				gstrWkrTitle = strSettingValue
			Case "SupervisorTitle"
				gstrSupTitle = strSettingValue
			Case "ManagerTitle"
				gstrMgrTitle = strSettingValue
			Case "OfficeTitle"
				gstrOffTitle = strSettingValue
			Case "DirectorTitle"
				gstrDirTitle = strSettingValue
			Case "ReviewerTitle" 
				gstrRvwTitle = strSettingValue
			Case "AuthorizeByTitle"
				gstrAuthTitle = strSettingValue
			Case "EvaluatorTitle"
				gstrEvaTitle = strSettingValue
			Case "Evaluation"
				gstrEvaluation = strSettingValue
            Case "WorkerRespDueDays"
				gstrWorkerRespDueDays = strSettingValue
			Case "AllowReviewEdit"
				gstrAllowReviewEdit = strSettingValue
			Case "ReviewClassTitle"
				gstrReviewClassTitle = strSettingValue
		    Case "UseFactorCodes"
		        gstrUseFactorCodes = strSettingValue
        End Select
        adRs.MoveNext
    Loop
    If gstrTitle = vbNullString Then
        gstrTitle = Trim(gstrLocationName & " " & gstrOrgName)
    End If

    'Customize IIS:
    Response.Buffer = True
    Server.ScriptTimeout = 120
End Function
%>
