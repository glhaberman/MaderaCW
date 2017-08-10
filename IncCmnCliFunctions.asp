<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncCmnCliFunctions.asp                                           '
' Purpose: This include file contains common client side functions.         '
'                                                                           '
'==========================================================================='
%>

<SCRIPT id=CommonClientScript language=vbscript>
<!--

Public Function Parse(strDelimitedList, strDelim, intPosition)
    Dim intPos
    Dim strWork
    Dim strValue
    Dim intCnt
    
    If intPosition = 0 Then
        Parse = ""
        Exit Function
    End If
    
    strWork = strDelimitedList
    
    intCnt = 0
    Do While strWork <> ""
        intCnt = intCnt + 1
        intPos = InStr(strWork, strDelim)
        If intPos = 0 Then
            If intCnt < intPosition Then
                Parse = ""
            Else
                Parse = strWork
            End If
            strWork = ""
        Else
            strValue = Left(strWork, intPos - 1)
            If intCnt = intPosition Then
                Parse = strValue
                Exit Function
            End If
            strWork = Mid(strWork, intPos + Len(strDelim))
        End If
    Loop
End Function

' Limts what type of characters can be typed in text boxes for various data types.
Function TextBoxOnKeyPress(intKeyAscii,strType)
	Select Case strType
		Case "N","Number","L","Long"
			Select Case	CLng(intKeyAscii)
				Case 48, 49, 50, 51,52, 53, 54, 55, 56, 57 ' Numbers 0 - 9
				Case Else
					window.event.keyCode = 0
			End Select
		Case "S","String","StringOther"
			Select Case	CLng(intKeyAscii)
				Case 39,34 ' Single Quote and Double Quotes cause problems.
					window.event.keyCode = 96 ' Uses [`] in place of [']
				Case 9,10,13,35,38,124 ' Ignore TAB, Carrige return and line feed
					window.event.keyCode = 0
			End Select
		Case "X" 'Letters, number and non-problamatic special characters
			Select Case	CLng(intKeyAscii)
				Case 39,34 ' Single Quote and Double Quotes cause problems.
					window.event.keyCode = 96 ' Uses [`] in place of [']
				Case 9,10,13,35,38,33,94,124,64 ' Ignore TAB, Carrige return and line feed, &, #, and "
					window.event.keyCode = 0
			End Select
		Case "C","Currency","Single"
			Select Case	CLng(intKeyAscii)
				Case 46, 48, 49, 50, 51, 52, 53, 54, 55, 56, 57 ' Numbers 0 - 9, [.]
				Case Else
					window.event.keyCode = 0
			End Select
		Case "D","Date"
			Select Case	CLng(intKeyAscii)
				Case 48, 49, 50, 51,52, 53, 54, 55, 56, 57, 47 ' Numbers 0 - 9, [/]
				Case Else
					window.event.keyCode = 0
			End Select
		Case "Time"
			Select Case	CLng(intKeyAscii)
				Case 48, 49, 50, 51,52, 53, 54, 55, 56, 57, 58 ' Numbers 0 - 9, [:]
				Case Else
					window.event.keyCode = 0
			End Select
	End Select
End Function

Function ValidDate(varDate)
    Dim blnErr
    Dim intPos1
    Dim intPos2
    Dim strMonth
    Dim strDay
    Dim strYear
    
    If Trim(varDate) = "" Then
        ValidDate = True
        Exit Function
    End If
    
    If Not IsDate(varDate) Then
        blnErr = True
    Else
        intPos1 = Instr(varDate, "/")
        intPos2 = Instr(intPos1 + 1, varDate, "/")
        
        If intPos1 = 0 Or intPos2 = 0 Then
            blnErr = True
        Else
            strMonth = Trim(Mid(varDate, 1, intPos1 - 1))
            strDay = Trim(Mid(varDate, intPos1 + 1, intPos2 - intPos1 - 1))
            strYear = Trim(Mid(varDate, intPos2 + 1))
            If strMonth = 0 Or strMonth > 12 Then
                blnErr = True
            ElseIf strDay = 0 Or strDay > 31 Then
                blnErr = True
            ElseIf Len(strYear) = 3 Or Len(strYear) > 4 Then
                blnErr = True
            Else
                If Len(strMonth) < 2 Then
                    strMonth = "0" & strMonth
                End If
                If Len(strDay) < 2 Then
                    strDay = "0" & strDay
                End If
                If Len(strYear) <= 2 Then
                    strYear = 2000 + CInt(strYear)
                End If
                varDate = strMonth & "/" & strDay & "/" & strYear
                If Not IsDate(varDate) Then
                    blnErr = True
                End If
            End If
        End If
    End If
    
    If blnErr Then
        ValidDate = False
    Else
        ValidDate = True
    End If
End Function

Sub SizeAndCenterWindow(intX, intY, blnCenter)
    Dim intScrLeft
    Dim intScrTop
    Dim intScrAvlX
    Dim intScrAvlY
    Dim blnMove

	' No longer resizing windows automatically.  To turn back on, remove Exit Sub below - 6/3/05 gh
	Exit Sub
	
    On Error Resume Next    
    window.resizeTo intX, intY
    
    intScrAvlX = window.screen.availWidth
    intScrAvlY = window.screen.availHeight

    If blnCenter Then
        'Center the window in the screen:
        window.moveTo (intScrAvlX - intX)/2, (intScrAvlY - intY)/2
    Else
        'Only move the window if it is not completely on screen:
        blnMove = False
        intScrLeft = Cint(window.screenLeft)
        intScrTop = Cint(window.screenTop - 30)
        If intScrLeft < 0 Then
            intScrLeft = 0
            blnMove = True
        End If
        If (intScrLeft + intX) > intScrAvlX Then
            intScrLeft = intScrLeft - Abs(intScrAvlX - (intScrLeft + intX))
            blnMove = True
        End If
        If intScrTop < 0 Then
            intScrTop = 0
            blnMove = True
        End If
        If (intScrTop + intY) > intScrAvlY Then
            intScrTop = intScrTop - Abs(intScrAVlY - (intScrTop + intY))
            blnMove = True
        End If 
        If blnMove Then
            window.moveTo intScrLeft, intScrTop
        End If
    End If

    On Error Goto 0
End Sub

Sub CheckForValidUser()
    'Validate the user when the windo loads - will navigate back 
    'to Logon form if the logon user id is not recognized:
    If Trim(Form.UserID.Value) = "" Then
        MsgBox "User not recognized.  Logon failed, please try again.", vbinformation, "Case Review Log On"
        window.navigate "Logon.asp"
    End If
End Sub

Function FormatDate(strDate)
    Dim strMonth
    Dim strDay
    Dim strYear
    
    If Not IsDate(strDate) Then
        Exit Function
    End If
    
    strMonth = Cstr(Month(strDate))
    strDay = Cstr(Day(strDate))
    strYear = Cstr(Year(strDate))
    
    Do While Len(strMonth) < 2
        strMonth = "0" & strMonth
    Loop
    Do While Len(strDay) < 2
        strDay = "0" & strDay
    Loop
    
    FormatDate = strMonth & "/" & strDay & "/" & strYear    
    
End Function

'------------------------------------------------------------------------------
' GetComboText:  This function is called to get the text of the currently 
'   selected option in a SELECT (combobox or listbox) element.  If there is
'   no current selection, the function returns an empty string.
'------------------------------------------------------------------------------
Function GetComboText(oCtl)
    Dim strVal

    strVal = ""
    
    On Error Resume Next
    If oCtl.selectedIndex > -1 Then
        strVal = oCtl.options(oCtl.selectedIndex).Text
    End If
    On Error Goto 0    
    
    GetComboText = strVal
End Function

Function GetComboTextByID(oCtl, strID)
    Dim strVal
    Dim intI

    strVal = ""
    
    For intI = 0 To oCtl.options.length - 1
        If CStr(oCtl.options(intI).Value) = Cstr(strID) Then
            strVal = oCtl.options(intI).Text
        End If
    Next
    
    GetComboTextByID = strVal
End Function

Sub CmnTxt_onfocus(txtCtl)
    txtCtl.Select
End Sub

Function Spacenb(intCnt)
    Dim intI
    Dim strSpace
    
    strSpace = ""
    For intI = 1 To intCnt
        strSpace = strSpace & "&nbsp"
    Next
    
    Spacenb = strSpace
End Function

Function LoadDictionaryObject(strDelimitedList)
    Dim strValue
    Dim intLast
    Dim intDelim
    Dim lngRecordID
    
    Set LoadDictionaryObject = CreateObject("Scripting.Dictionary")
    If Len(strDelimitedList) <= 1 Then Exit Function
	intLast = 1
	
	' First item in delimited list will be the key and the entire sting will be the item
	Do While True
		intDelim = Instr(intLast + 1,strDelimitedList,"|")
		strValue = Mid(strDelimitedList, intLast + 1, intDelim - (intLast + 1))
        lngRecordID = Parse(strValue,"^",1)
		LoadDictionaryObject.Add CLng(lngRecordID),strValue
		intLast = intDelim
		If intLast = Len(strDelimitedList) Then Exit Do		
	Loop
End Function
-->
</script>
