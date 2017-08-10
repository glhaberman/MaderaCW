<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncValidUser.asp                                                 '
' Purpose: This include file contains common code for validating the user   '
'          ID and password from the logon screen.                           '
'                                                                           '
'==========================================================================='

Dim madoRsUser      'Recordset for user lookup.
Dim mintChrPos      'Loop counter for moving through user name string.
Dim blnAdminLogon   'An admin user is logging in under another id.

gstrUserID = ""
gblnUserAdmin = Null
gblnUserQA = Null

If gblnUseLogon Then
    'Validate the user id and password passed to the page:
    If Trim(Request.Form("UserID")) = "" Or Trim(Request.Form("Password")) = "" Then
        Response.redirect "Logon.asp"
    End If
            
    Set gadoCmd = Server.CreateObject("ADODB.Command")
    With gadoCmd
        .ActiveConnection = gadoCon
        .CommandType = adCmdStoredProc
        .CommandText = "spCheckUser"
        gstrUserID = Request.Form("UserID")
        If Request.Form("CalledFrom") = "Logon" Then
            'Arriving from the Logon page:
            gstrPassword = Encrypt(LCase(Request.Form("Password")), UCase(gstrUserID))
        Else
            'Arriving from all other pages, the password is already encrypted:
            gstrPassword = Request.Form("Password")
        End If
        .Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 20, gstrUserID)
        .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 60, gstrPassword)
        'Call ShowCmdParms(gadoCmd) '***DEBUG
    End With
    'Open a recordset from the query:
    Set madoRsUser = Server.CreateObject("ADODB.Recordset") 
    Call madoRsUser.Open(gadoCmd, , adOpenForwardOnly, adLockReadOnly)

    gstrUserID = ""
    gstrPassword = Null
Else
    ' Do not validate Windows user ID if there is a UserID passed in.
    If Request.Form("UserID") = "" Then
	
	    'Validate the Windows user id:
	    gstrUserID = Request.ServerVariables("AUTH_USER")
	    'Remove domain name:
	    If Instr(gstrUserID, "\") > 0 Then
		    For mintChrPos = Len(gstrUserID) To 1 Step -1
			    If Mid(gstrUserID, mintChrPos, 1) = "\" Then
				    mintChrPos = mintChrPos + 1
				    Exit For
			    End If
		    Next
		    gstrUserID = Mid(gstrUserID, mintChrPos)
	    End If
	    'Abort the page if the IIS authentication is not set properly:
	    'If gstrUserID = "" Or Instr("negotiatentlm", LCase(Request.ServerVariables("AUTH_TYPE"))) = 0 Then
		'    Response.write "Case Review System<br><br>** Unable to determine network user ID.<br>** Logon failed.<br>"
		'    Response.Write "AUTH_TYPE = " & Request.ServerVariables("AUTH_TYPE")
		'    Response.End
	    'End If
    Else
	    ' If a UserID is passed in, set the global variable here
	    gstrUserID = Request.Form("UserID")
    End If

    'System Admin user is switching to another logon:
    If Instr(Request.Form("UserID"), "**ADMIN**") > 0 Then 
        gstrUserID = Mid(Request.Form("UserID"), 10)
        gstrPassword = Request.Form("Password")
        blnAdminLogon = True
    End If

    'Lookup the Windows ID in the case review system:
    Set gadoCmd = Server.CreateObject("ADODB.Command")
    With gadoCmd
        .ActiveConnection = gadoCon
        .CommandType = adCmdStoredProc
        .CommandText = "spCheckUser"
        gstrUserID = gstrUserID
        .Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 20, gstrUserID)
        .Parameters.Append .CreateParameter("@Password", adVarChar, adParamInput, 60, chr(9) & "^INTEGRATED^")
    End With
    Set madoRsUser = Server.CreateObject("ADODB.Recordset") 
    Call madoRsUser.Open(gadoCmd, , adOpenForwardOnly, adLockReadOnly)
End If  '[gblnUseLogon]

gstrUserID = ""
gstrPassword = Null
Do While Not madoRsUser.EOF
    'If the user ID was valid, the ID and staff information are
    'stored in variables used to generate the clientside page
    'load.  If the values are not set here (i.e. invalid user)
    'then the page won't load:
    gstrUserID = madoRsUser.Fields("logUserLogin")              'User Login of current user
    gstrPassword = madoRsUser.Fields("logPassword")             'Password
    gstrOptions = madoRsUser.Fields("Options")                  'String of security options for login
    gstrRoles = madoRsUser.Fields("Roles")                      'String of security roles for login
    gstrUserName = madoRsUser.Fields("UserName")                'The full employee name of login
    glngAliasPosID = madoRsUser.Fields("logAliasID")            'Aliased position ID, or 0 if none
    madoRsUser.MoveNext
Loop

madoRsUser.Close
Set gadoCmd = Nothing

gblnUserQA = False
gblnUserAdmin = False

'1 -- Re-Reviewer
'2 -- Reviewer
'3 -- Administrator
'4 -- System Administrator
'5 -- Reports
If Instr(gstrRoles, "[3]") > 0 Then
	gblnUserAdmin = True 
End IF
If Instr(gstrRoles, "[4]") > 0 Then
	gblnUserQA = True  
End IF

If gstrUserID <> "" And ReqForm("ProgramsSelected") <> "" Then
    'Save the current selected programs:
    Set gadoCmd = GetAdoCmd("spProfileSettingUpd")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, gstrUserID
        AddParmIn gadoCmd, "@SettingName", adVarChar, 50, "ProgramsSelected"
        AddParmIn gadoCmd, "@SettingValue", adVarChar, 255, ReqForm("ProgramsSelected")
        gadoCmd.Execute
    Set gadoCmd = Nothing
End If

Function Encrypt(strEncryptString, strKey)
    Dim intPos
    Dim intTmp
    Dim X1
    Dim strX1
    Dim intKey
    Dim G
    Dim strEncrypted

    'Simple string encryption/decryption:
    For intPos = 1 To Len(strKey)
        intTmp = Asc(Mid(strKey, intPos, 1))
        X1 = X1 + intTmp
    Next

    X1 = Int((X1 * 0.1) / 6)
    intKey = X1
    G = 0

    For intPos = 1 To Len(strEncryptString)
        intTmp = Asc(Mid(strEncryptString, intPos, 1))
        
        G = G + 1
        If G = 6 Then G = 0
        
        X1 = 0

        If G = 0 Then X1 = intTmp - (intKey - 2)
        If G = 1 Then X1 = intTmp + (intKey - 5)
        If G = 2 Then X1 = intTmp - (intKey - 4)
        If G = 3 Then X1 = intTmp + (intKey - 2)
        If G = 4 Then X1 = intTmp - (intKey - 3)
        If G = 5 Then X1 = intTmp + (intKey - 5)
        X1 = X1 + G
        
        strX1 = CStr(X1)
        
        Do While Len(strX1) < 3
            strX1 = "0" & strX1
        Loop
        
        strEncrypted = strEncrypted & strX1
    Next

    Encrypt = strEncrypted
End Function

Function Decrypt(strEncryptString, strKey)
    Dim intPos
    Dim intTmp
    Dim X1
    Dim intKey
    Dim G
    Dim strDecrypted

    'Simple string encryption/decryption:
    For intPos = 1 To Len(strKey)
       intTmp = Asc(Mid(strKey, intPos, 1))
       X1 = X1 + intTmp
    Next

    X1 = Int((X1 * 0.1) / 6)
    intKey = X1
    G = 0
    
    For intPos = 1 To Len(strEncryptString) Step 3
        intTmp = Mid(strEncryptString, intPos, 3)
        
        G = G + 1
        If G = 6 Then G = 0
        
        X1 = 0

        If G = 0 Then X1 = intTmp + (intKey - 2)
        If G = 1 Then X1 = intTmp - (intKey - 5)
        If G = 2 Then X1 = intTmp + (intKey - 4)
        If G = 3 Then X1 = intTmp - (intKey - 2)
        If G = 4 Then X1 = intTmp + (intKey - 3)
        If G = 5 Then X1 = intTmp - (intKey - 5)
        X1 = X1 - G
        
        strDecrypted = strDecrypted & Chr(X1)
    Next

    Decrypt = strDecrypted
End Function

%>