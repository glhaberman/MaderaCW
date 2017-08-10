<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncInteliType.asp                                                '
' Purpose: Contains the declarations, event procedures, and functions for   '
'          implementing the inteli-type feature on a <SELECT> element.      '
'                                                                           '
' Instructions:                                                             '
'   -Include this file in the page containing the <SELECT> element.         '
'   -Create OnClick, OnFocus, OnBlur, OnKeyPress, and OnKeyDown event code  '
'   for the <SELECT> element in the parent page if they do not exist.       '
'   -In each of those event procedures on the parent page, call the         '
'   corresponding inteli-type event here.  Note that the OnKeyPress and     '
'   OnKeyDown require a reference to the control triggering the event.      '
'                                                                           '
'   For example:                                                            '
'       Sub cboReviewType_onclick()                                         '
'           Call cboInteliOnClick()                                         '
'       End Sub                                                             '
'       Sub cboReviewType_onfocus()                                         '
'           Call cboInteliOnFocus()                                         '
'       End Sub                                                             '
'       Sub cboReviewType_onblur()                                          '
'           Call cboInteliOnBlur()                                          '
'       End Sub                                                             '
'       Sub cboReviewType_onkeypress()                                      '
'           Call cboInteliOnKeyPress(cboReviewType)                         '
'       End Sub                                                             '
'       Sub cboReviewType_onkeydown()                                       '
'           Call cboInteliOnKeyDown(cboReviewType)                          '
'       End Sub                                                             '
'==========================================================================='
%>

<SCRIPT id=InteliType language=vbscript>
<!--
Dim mstrValue       'Holds the characters typed in combo after receiving focus.
Dim mdteLastKeyTime 'Last time key was pressed in combobox.  Intelitype resets
                    'after a number of seconds.

Sub cboInteliOnclick()
    'Clear the keystroke holder whenever combo is clicked by mouse:
    mstrValue = ""
End Sub

Sub cboInteliOnfocus()
    'Clear the keystroke holder when control receives focus:
    mstrValue = ""
    'Initialize the keystroke time:
    mdteLastKeyTime = Now
End Sub

Sub cboInteliOnBlur()
    'Clear the keystroke holder when leaving the control:
    mstrValue = ""
End Sub

Sub cboInteliOnKeyPress(oListCtl)
    'If 5 seconds has elapsed since last keystroke, reset the keystroke holder:
    If DateDiff("s", mdteLastKeyTime, Now) > 5 Then
        'Reset
        mstrValue = ""
    End If
    
    'Add the current key pressed to the keystroke holder:
    mstrValue = mstrValue & chr(Window.event.keyCode)
    'Cancel the event to prevent the native matching:
    Window.event.returnValue = False
    'Do our custom matching, searching the list for what has been typed:
    Call FindInList(oListCtl, mstrValue)
    
    'Set the time of this keypress:
    mdteLastKeyTime = Now
End Sub

Sub cboInteliOnKeyDown(oListCtl)
    'Watch for backspace or delete key.
    If Window.event.keyCode = 8 Then 'Backspace
        'If backspace, then cancel the event to prevent browser from boing back,
        'and delete one character from the end of the keystroke holder:
        Window.event.returnValue = False
        If Len(mstrValue) > 0 Then
            mstrValue = Left(mstrValue, Len(mstrValue) - 1)
        End If
        'Research for the value after character was deleted:
        Call FindInList(oListCtl, mstrValue)
    ElseIf Window.event.keyCode = 46 Then
        'If delete, cancel the event and clear the keystroke holder:
        Window.event.returnValue = False
        mstrValue = ""
        'If the list has an entry for "blank", it will be selected,
        'otherwise the list will just be left at the last entry found.
        Call FindInList(oListCtl, mstrValue)
    ElseIf Window.event.keyCode = 40 Or Window.event.keyCode = 38 Then
        'Up or down arrows reset the keystroke holder:
        mstrValue = ""
    End If
End Sub

Sub FindInList(oListCtl, strVal)
    Dim intI
    'Search for the string in "strVal" in the SELECT's option list:
    For intI = 0 To oListCtl.Options.length - 1
        If UCase(strVal) = UCase(Left(oListCtl.options(intI).text, Len(strVal))) Then
            oListCtl.selectedIndex = intI
            If oListCtl.onchange = "" Or IsNull(oListCtl.onchange) Then
                On Error Resume Next
                oListCtl.onchange = GetRef(oListCtl.ID & "_onchange")
                On Error Goto 0
            End If
            If Not IsNull(oListCtl.onchange) Then
                Call oListCtl.onchange
            End If
            Exit For
        End If
    Next
End Sub
-->
</script>
