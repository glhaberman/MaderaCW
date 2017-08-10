<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncSvrFunctions.asp                                              '
' Purpose: This include file contains any common server side functions.     '
'                                                                           '
'==========================================================================='
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

Function ClearScript(strText)
    If IsNull(strText) Then
        ClearScript = Null
    Else 'If strText = "" 
        ClearScript = Server.HTMLEncode(strText)
    End If
End Function

Function GetTabIndex()
    mlngTabIndex = mlngTabIndex + 1
    GetTabIndex = mlngTabIndex
End Function

Function GetAdoCmd(strStoredProc)
    Dim adCmd
    
    Set adCmd = Server.CreateObject("ADODB.Command")
    With adCmd
        .ActiveConnection = gadoCon
        .CommandType = adCmdStoredProc
        .CommandText = strStoredProc
        .CommandTimeout = 180
    End With    
    Set GetAdoCmd = adCmd
    Set adCmd = Nothing    
End Function

Function GetAdoRs(adCmd)
    Dim rs
    
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
    
    Set GetAdoRs = rs
    Set rs = Nothing
End Function

Sub AddParmIn(adCmd, strName, intType, intLen, vntValue)
    On Error Resume Next
    With adCmd
        .Parameters.Append .CreateParameter(strName, intType, adParamInput, intLen, vntValue)
        'If Err.number > 0 Then
		'	response.Write strName & "=" & vntValue
		'	response.End
		'End IF
    End With
    On Error Goto 0
End Sub

Sub AddParmOut(adCmd, strName, intType, intLen)
    With adCmd
        .Parameters.Append .CreateParameter(strName, intType, adParamOutput, intLen, NULL)
    End With
End Sub

Function ReqForm(strField)
    ReqForm = Trim(Request.Form(strField))
End Function

Function ReqIsBlank(strField)
    Dim strTmp
    
    strTmp = ReqForm(strField)
    ReqIsBlank = IsBlank(strTmp)
End Function

Function ReqZeroToNull(strField)
    Dim strTmp
    
    strTmp = ReqIsNumeric(strField)
    ReqZeroToNull = ZeroToNull(strTmp)
End Function

Function ZeroToNull(strValue)
    Dim strTmp
    If IsNumeric(strValue) Then
        strTmp = strValue
    Else
        strTmp = 0
    End If
    If strTmp = 0 Then
        ZeroToNull = Null
    Else
        ZeroToNull = strTmp
    End If
End Function

Function IsBlank(strValue)
    If Trim(strValue) = "" Then
        IsBlank = NULL
    Else
        IsBlank = strValue
    End If
End Function

Function ReqIsNumeric(strField)
    If IsNumeric(ReqForm(strField)) Then
        ReqIsNumeric = ReqForm(strField)
    Else
        ReqIsNumeric = 0
    End If
End Function

Function ReqIsDate(strField)
    If IsDate(ReqForm(strField)) Then
        ReqIsDate = ReqForm(strField)
    Else
        ReqIsDate = NULL
    End If
End Function

Function OpenConnection(strServer, strDatabase, strUser, strPassword)
    Dim strCnn
    If Len(Trim(strUser)) = 0 Then
        'Assume SQL Windows authentication is being used:
        strCnn = "driver={SQL Server};server=" & strServer & ";database=" & strDb
    Else
        'Attempt connection using the user and password information:
        strCnn = "driver={SQL Server};server=" & strServer & ";uid=" & strSqlUser & ";pwd=" & strSqlPW & ";database=" & strDb
    End If

    'Create and open an ADO connection to the database:
    On Error Resume Next
    Set gadoCon = Server.CreateObject("ADODB.Connection") 
    gadoCon.CursorLocation = adUseClient
    gadoCon.Open strCnn
    If err.number <> 0 Then
        Response.Write "A problem occurred connecting to the Case Review database '" & strDb & "':<br><br>"
        Response.Write err.number & "<br>" & err.Description & "<br>"
        Response.Write gadoCon.Errors(0).NativeError & "<br>" & gadoCon.Errors(0).Description & "<br>"
        Response.End
    End If
    On Error Goto 0
End Function

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
%>
<!--#include file="IncGetAppSetting.asp"-->
