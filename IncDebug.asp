<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncDebug.asp                                                     '
' Purpose: This include is used on all pages in the application.  If the    '
'          debugging flag gblnDebugOn is set to True in IncGlobals, then    '
'          the functions (both server and client) are enabled.  If the flag '
'          is False, the functions will do nothing.  Ideally, this include  '
'          and all references to it would be cleaned out in the final       '
'          stages of development.                                           '
'==========================================================================='

'Server side debugging functions:
Public Function ShowCmdParms(adCmd)
    Dim intCnt

    If gblnDebugOn Then
        Response.Write vbCrLf & "***Debug***<br>" & vbCrLf
        Response.Write adCmd.CommandText & vbCrLf & "<br>-------------------------------------------<br>" & vbCrlf
        For intCnt = 0 To adCmd.Parameters.count - 1
            Response.Write adCmd.Parameters.Item(intCnt).Name & "="
            If IsNull(adCmd.Parameters.Item(intCnt).Value) Then
                Response.Write "NULL"
            Else
                Select Case TypeName(adCmd.Parameters.Item(intCnt).Value)
                    Case "String", "Date"
                        Response.Write "'" & adCmd.Parameters.Item(intCnt).Value & "'"
                    Case Else
                        Response.Write adCmd.Parameters.Item(intCnt).Value
                End Select
                 typename(adCmd.Parameters.Item(intCnt).Value) & ")"
            End If
            Response.Write "<br>" & vbCrLf
        Next
        Response.End
    End If
End Function

Public Sub DebugMsg(strMsg)
    If gblnDebugOn Then
        Response.Write "MsgBox strMsg, vbInformation, ""***DEBUG***""" & vbCrLf
    End If
End Sub
%>