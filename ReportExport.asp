<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ReportExport.asp                                                '
'  Purpose: This page is used to execute the stored procedure that generated'
'           the report that called this page.  The results are copied to a  '
'           TAB delimeted clipboard.                                        '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd
Dim adRs
Dim adRsChild
Dim mstrStoredProcedureName
Dim mstrReportName
Dim mstrParameters
Dim mstrParameter
Dim mstrClipboard
Dim intI
Dim intJ
Dim intK
Dim intL
Dim maFields()
Dim mstrMultiValues
Dim mstrDelim
Dim mstrRptDetailLayout
Dim mstrParentClip
Dim mstrChildClip
Dim mstrWorkerAuthBy

mstrStoredProcedureName = Request.QueryString("SPName")
mstrReportName = Request.QueryString("RName")
mstrParameters = Request.QueryString("Parameters")

mstrRptDetailLayout = GetAppSetting("RptDetailLayout")

' Get Report Export Column Names
Set adCmd = GetAdoCmd("spGetReportExportNames")
adCmd.CommandTimeout = 180
AddParmIn adCmd, "@ReportASP", adVarChar, 50, mstrReportName
'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
intI = 0
Do While Not adRs.Eof
    ReDim Preserve maFields(3,intI)
    maFields(0,intI) = adRs("rpxSPFieldName").value
    maFields(1,intI) = adRs("rpxExportFieldName").value
    If Len(maFields(1,intI)) = 0 Then maFields(1,intI) = "NOPARSE"
    maFields(2,intI) = adRs("rpxParse").value
    maFields(3,intI) = adRs("rpxParseCols").value
    intI = intI + 1   
    adRs.MoveNext
Loop
' Load recordset
Set adCmd = GetAdoCmd(mstrStoredProcedureName)
adCmd.CommandTimeout = 180
mstrParameters = mstrParameters & "**End**^^^"
intI = 1
Do While True
    mstrParameter = Parse(mstrParameters,"|",intI)
    If Parse(mstrParameter,"^",1) = "**End**" Then Exit Do
    
    If Parse(mstrParameter,"^",1) = "@ReportingMode" Then
        If Parse(mstrParameter,"^",4) = "1" Then
            mstrWorkerAuthBy = "Authorized By"
        Else
            mstrWorkerAuthBy = "Worker"
        End If    
    End If
    If Parse(mstrParameter,"^",4) = "" Then
        AddParmIn adCmd, Parse(mstrParameter,"^",1), Parse(mstrParameter,"^",2), Parse(mstrParameter,"^",3), Null
    Else
        AddParmIn adCmd, Parse(mstrParameter,"^",1), Parse(mstrParameter,"^",2), Parse(mstrParameter,"^",3), Parse(mstrParameter,"^",4)
    End If
    intI = intI + 1
Loop
Set adRs = GetAdoRs(adCmd)

If adRs.BOF And adRs.EOF Then
    mstrClipboard = "* No reviews matched the report criteria *"
Else
    mstrClipboard = ""
    ' Setup column headings
    For intI = 0 To adRs.Fields.Count - 1
        For intJ = 0 To UBound(maFields,2)
            If maFields(0,intJ) = adRs.Fields(intI).Name Then
                If maFields(2,intJ) = "NOPARSE" Then
                    mstrClipboard = mstrClipboard & maFields(1,intJ) & Chr(9)
                Else
                    ' Recordset column contains multiple values, with column headers delimited with ^
                    mstrMultiValues = maFields(1,intJ)
                    If mstrMultiValues = "Worker Name^Worker Number" Then
                        ' Update column headers for Worker or Auth By to what was clicked on Reports page
                        mstrMultiValues = Replace(mstrMultiValues, "Worker", mstrWorkerAuthBy)
                    End If
                    mstrMultiValues = mstrMultiValues & "^ZZZ"
                    intK = 1
                    Do While intK <= maFields(3,intJ) + 1 ' Allow 1 for the ZZZ
                        If Parse(mstrMultiValues,"^",intK) = "ZZZ" Then 
                            If intK <= maFields(3,intJ) Then ' If # of columns is less than designated, insert null strings
                                For intL = intK To maFields(3,intJ)
                                    mstrClipboard = mstrClipboard & "unknown" & Chr(9)
                                Next
                            End If
                            Exit Do
                        End If
                        mstrClipboard = mstrClipboard & Parse(mstrMultiValues,"^",intK) & Chr(9)
                        intK = intK + 1
                    Loop
                End If
                Exit For
            End If
        Next
    Next
End If

Select Case mstrReportName
    Case "RptDetail"
        Call RptDetailExport()
    Case "RptMitigation"
        Call RptMitigationExport()
    Case Else
        Call GenericExport()
End Select

Sub GenericExport()
    If mstrClipboard <> "* No reviews matched the report criteria *" Then
        ' Copy recordset values to clipboard
        Do While Not adRs.EOF
            mstrClipboard = mstrClipboard & "**vbCrLf**"
            For intI = 0 To adRs.Fields.Count - 1
                For intJ = 0 To UBound(maFields,2)
                    If maFields(0,intJ) = adRs.Fields(intI).Name Then
                        If maFields(2,intJ) <> "NOPARSE" Then
                            ' Recordset column contains multiple values, with column headers delimited with aFields(2,intJ)
                            mstrMultiValues = adRs.Fields(intI).Value & maFields(2,intJ) & "ZZZ"
                            intK = 1
                            Do While intK <= maFields(3,intJ) + 1 ' Allow 1 for the ZZZ
                                If Parse(mstrMultiValues,maFields(2,intJ),intK) = "ZZZ" Then 
                                    If intK <= maFields(3,intJ) Then ' If # of columns is less than designated, insert null strings
                                        For intL = intK To maFields(3,intJ)
                                            If InStr(adRs.Fields(intI).Value,"Vacant") > 0 Then
                                                ' If # of columns is less than designated and the word "Vacant" appears
                                                ' in another column, assume PositionID belongs in empty column.
                                                mstrClipboard = mstrClipboard & trim(Mid(adRs.Fields(intI).Value,7,Len(adRs.Fields(intI).Value)-6)) & Chr(9)
                                            Else
                                                mstrClipboard = mstrClipboard & "" & Chr(9)
                                            End If
                                        Next
                                    End If
                                    Exit Do
                                End If
                                mstrClipboard = mstrClipboard & Trim(Parse(mstrMultiValues,maFields(2,intJ),intK)) & Chr(9)
                                intK = intK + 1
                            Loop
                        Else
                            mstrClipboard = mstrClipboard & adRs.Fields(intI).Value & Chr(9)
                        End If
                        Exit For
                    End If
                Next
            Next
            adRs.MoveNext
        Loop
    End If
    mstrClipboard = Replace(mstrClipboard,"""","#dq#")
End Sub

Sub RptDetailExport()
    Dim strShowDetail
    
    strShowDetail = Request.QueryString("ShowDetail")

    If mstrClipboard <> "* No reviews matched the report criteria *" Then
        If strShowDetail = "Y" Then
            If mstrRptDetailLayout = "Single Line" Then
                mstrClipboard = mstrClipboard & "Program" & Chr(9) & _
                    "Element" & Chr(9) & _
                    "Element Status" & Chr(9) & _
                    "Factor 1" & Chr(9) & _
                    "Factor 2" & Chr(9) & _
                    "Factor 3" & Chr(9) & _
                    "Case Action" & Chr(9) & _
                    "Review Type" & Chr(9) & "**vbCrLf**"
            Else ' Header / Line Item
                mstrClipboard = mstrClipboard & "**vbCrLf**"
                mstrClipboard = mstrClipboard & Chr(9) & "Program" & Chr(9) & _
                    "Element" & Chr(9) & _
                    "Element Status" & Chr(9) & _
                    "Factor 1" & Chr(9) & _
                    "Factor 2" & Chr(9) & _
                    "Factor 3" & Chr(9) & _
                    "Case Action" & Chr(9) & _
                    "Review Type" & Chr(9) & "**vbCrLf**"
            End If
        End If
        
        ' Copy recordset values to clipboard
        Do While Not adRs.EOF
            If strShowDetail <> "Y" Then
                mstrClipboard = mstrClipboard & "**vbCrLf**"
            End If
            mstrParentClip = ""
            mstrChildClip = ""
            For intI = 0 To adRs.Fields.Count - 1
                For intJ = 0 To UBound(maFields,2)
                    If maFields(0,intJ) = adRs.Fields(intI).Name Then
                        If maFields(2,intJ) <> "NOPARSE" Then
                            ' Recordset column contains multiple values, with column headers delimited with aFields(2,intJ)
                            mstrMultiValues = adRs.Fields(intI).Value & maFields(2,intJ) & "ZZZ"
                            intK = 1
                            Do While intK <= maFields(3,intJ) + 1 ' Allow 1 for the ZZZ
                                If Parse(mstrMultiValues,maFields(2,intJ),intK) = "ZZZ" Then 
                                    If intK <= maFields(3,intJ) Then ' If # of columns is less than designated, insert null strings
                                        For intL = intK To maFields(3,intJ)
                                            If InStr(adRs.Fields(intI).Value,"Vacant") > 0 Then
                                                ' If # of columns is less than designated and the word "Vacant" appears
                                                ' in another column, assume PositionID belongs in empty column.
                                                mstrParentClip = mstrParentClip & Trim(Mid(adRs.Fields(intI).Value,7,Len(adRs.Fields(intI).Value)-6)) & Chr(9)
                                            Else
                                                mstrParentClip = mstrParentClip & "" & Chr(9)
                                            End If
                                        Next
                                    End If
                                    Exit Do
                                End If
                                mstrParentClip = mstrParentClip & Trim(Parse(mstrMultiValues,maFields(2,intJ),intK)) & Chr(9)
                                intK = intK + 1
                            Loop
                        Else
                            mstrParentClip = mstrParentClip & adRs.Fields(intI).Value & Chr(9)
                        End If
                        Exit For
                    End If
                Next
            Next
            If strShowDetail = "Y" Then
                ' Insert child records here
                Set adCmd = GetAdoCmd("spGetReviewElements")
                    AddParmIn adCmd, "@rvwID", adInteger, 0, adRs.Fields("rvwID").value
                    'Call ShowCmdParms(adCmd) '***DEBUG
                Set adRsChild = GetAdoRs(adCmd)
                If mstrRptDetailLayout <> "Single Line" Then
                    mstrClipboard = mstrClipboard & mstrParentClip & "**vbCrLf**"
                End If
                Do While Not adRsChild.EOF
                    mstrChildClip = ""
                    For intI = 0 To adRsChild.Fields.Count - 1
                        mstrChildClip = mstrChildClip & adRsChild.Fields(intI) & Chr(9)
                    Next
                    If mstrRptDetailLayout <> "Single Line" Then
                        mstrClipboard = mstrClipboard & Chr(9) & mstrChildClip & "**vbCrLf**"
                    Else
                        mstrClipboard = mstrClipboard & mstrParentClip & mstrChildClip & "**vbCrLf**"
                    End If
                    adRsChild.MoveNext
                Loop
                adRsChild.Close
            Else
                mstrClipboard = mstrClipboard & mstrParentClip
            End If            
            adRs.MoveNext
        Loop
    End If
    mstrClipboard = Replace(mstrClipboard,"""","#dq#")
End Sub

Sub RptMitigationExport()
    Dim strShowDetail
    Dim intRSCounter
    Dim blnFirstRecord
    
    blnFirstRecord = True
    
    strShowDetail = Request.QueryString("ShowDetail")
    If mstrClipboard <> "* No reviews matched the report criteria *" Then
        ' Copy recordset values to clipboard
        mstrClipboard = mstrClipboard & "**vbCrLf**"
        intRSCounter = 0
        Do While Not adRs Is Nothing
            If adRs.State = adStateOpen Then
                intRSCounter = intRSCounter + 1
                Do While Not adRs.EOF
                    For intI = 0 To adRs.Fields.Count - 1
                        If intRSCounter = 1 Then
                            For intJ = 0 To UBound(maFields,2)
                                If maFields(0,intJ) = adRs.Fields(intI).Name Then
                                    If maFields(2,intJ) <> "NOPARSE" Then
                                        ' Recordset column contains multiple values, with column headers delimited with aFields(2,intJ)
                                        mstrMultiValues = adRs.Fields(intI).Value & maFields(2,intJ) & "ZZZ"
                                        intK = 1
                                        Do While intK <= maFields(3,intJ) + 1 ' Allow 1 for the ZZZ
                                            If Parse(mstrMultiValues,maFields(2,intJ),intK) = "ZZZ" Then 
                                                If intK <= maFields(3,intJ) Then ' If # of columns is less than designated, insert null strings
                                                    For intL = intK To maFields(3,intJ)
                                                        If InStr(adRs.Fields(intI).Value,"Vacant") > 0 Then
                                                            ' If # of columns is less than designated and the word "Vacant" appears
                                                            ' in another column, assume PositionID belongs in empty column.
                                                            mstrClipboard = mstrClipboard & Trim(Mid(adRs.Fields(intI).Value,7,Len(adRs.Fields(intI).Value)-6)) & Chr(9)
                                                        Else
                                                            mstrClipboard = mstrClipboard & "" & Chr(9)
                                                        End If
                                                    Next
                                                End If
                                                Exit Do
                                            End If
                                            mstrClipboard = mstrClipboard & Trim(Parse(mstrMultiValues,maFields(2,intJ),intK)) & Chr(9)
                                            intK = intK + 1
                                        Loop
                                    Else
                                        mstrClipboard = mstrClipboard & adRs.Fields(intI).Value & Chr(9)
                                    End If
                                    Exit For
                                End If
                            Next
                        Else
                            If blnFirstRecord = True Then
                                If strShowDetail = "Y" Then
                                    mstrClipboard = mstrClipboard & "**vbCrLf**" & "Mitigated Cases Details" & "**vbCrLf**"
                                    mstrClipboard = mstrClipboard & _
                                        "Case Number" & Chr(9) & _
                                        "Case Name" & Chr(9) & "**vbCrLf**"
                                Else
                                    ' If ShowDetails is N, no need to continue with second recordset
                                    Exit Do
                                End If
                                blnFirstRecord = False
                            End If
                            mstrClipboard = mstrClipboard & adRs.Fields(intI).Value & Chr(9)
                        End If
                    Next
                    mstrClipboard = mstrClipboard & "**vbCrLf**"
                    adRs.MoveNext
                Loop
            End If
            'mstrClipboard = mstrClipboard & "**vbCrLf**" & mstrParentClip & "**vbCrLf**"
            Set adRs = adRs.NextRecordset
        Loop
    End If
    mstrClipboard = Replace(mstrClipboard,"""","#dq#")
End Sub

%>
<HTML><HEAD>
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
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
    Call window.clipboardData.setData("Text", Replace(Replace("<% = mstrClipboard %>","**vbCrLf**",vbCrLf),"#dq#",""""))
    window.close
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:#white">
<BR><BR><BR>
<DIV id=divDisplay>
    Copying data to clipboard, please wait...
</DIV>
</BODY>
</HTML>
