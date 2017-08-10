<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ElementSort.asp                                           '
'  Purpose: This page is used to edit benefit fields and Submit fields.     '
' Includes:                                                                 '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd, madoRs
Dim intElmID1, intElmID2
Dim strAction, mdctReturn
Dim intProgramID, intTabID, oReturn
Dim mstrWaitMessage

strAction = Request.QueryString("Action")
mstrWaitMessage = "Accessing Database...Please Wait..."

Set adCmd = Server.CreateObject("ADODB.Command")
Set madoRs = Server.CreateObject("ADODB.Recordset")
Set mdctReturn = Server.CreateObject("Scripting.Dictionary")
Select Case strAction
    Case "Sort"
        intElmID1 = Request.QueryString("ElmID1")
        intElmID2 = Request.QueryString("ElmID2")
        Set adCmd = GetAdoCmd("spElementSort")
            AddParmIn adCmd, "@ElmID1", adInteger, 0, intElmID1
            AddParmIn adCmd, "@ElmID2", adInteger, 0, intElmID2
            'Call ShowCmdParms(adCmd) '***DEBUG
        adCmd.Execute
        mdctReturn.Add 0,""
    Case "GetActions"
        Set adCmd = GetAdoCmd("spGetElements")
            AddParmIn adCmd, "@ProgramID", adInteger, 0, Request.QueryString("ProgramID")
            AddParmIn adCmd, "@TypeID", adInteger, 0, Request.QueryString("TabID")
            madoRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
        Set adCmd = Nothing
        Do While Not madoRs.EOF
            mdctReturn.Add CLng(madoRs.Fields("elmID").Value), madoRs.Fields("elmProgramID").Value & "^" & madoRs.Fields("elmShortName").Value 
            madoRs.MoveNext
        Loop
    Case "FactorLinks"
        Set adCmd = GetAdoCmd("spFactorLinks")
            AddParmIn adCmd, "@FactorID", adInteger, 0, Request.QueryString("FactorID")
            madoRs.Open adCmd, , adOpenForwardOnly, adLockReadOnly
        Set adCmd = Nothing
        If madoRs.RecordCount > 0 Then
            madoRs.Sort = "Element, Program"
        End If
        Do While Not madoRs.EOF
            mdctReturn.Add CLng(madoRs.Fields("ElementID").Value), madoRs.Fields("TabName").Value & "^" & _
                madoRs.Fields("Program").Value & "^" & _
                madoRs.Fields("Element").Value & "^" & _
                madoRs.Fields("LinkID").Value & "^" & _
                madoRs.Fields("LinkEndDate").Value & "^" & _
                madoRs.Fields("LinkLastDate").Value
            madoRs.MoveNext
        Loop
End Select


Response.ExpiresAbsolute = Now - 5
%>
<HTML><HEAD>
<TITLE>Element Sort</TITLE>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit
Dim mdctReturn

Sub window_onload()
    Set mdctReturn = CreateObject("Scripting.Dictionary")
    <%
    For Each oReturn In mdctReturn
        Response.Write "mdctReturn.Add CLng(" & oReturn & "), """ & mdctReturn(oReturn) & """" & vbCrLf
    Next
    %>
    window.returnvalue = mdctReturn
    window.close
End Sub

</SCRIPT>
<BODY>
<BR><BR><%=mstrWaitMessage%>
</BODY>
</HTML>
