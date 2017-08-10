<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ActivityAudit.asp                                               '
'  Purpose: This page is used to insert activity audit actions.             '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd
Dim adRs
Dim strAction
Dim strUserID
Dim strDescription
Dim strComments
Dim strTable
Dim lngRecordID
Dim mstrDetails
Dim mstrChangeDate
Dim mstrAudit
Dim mstrComments

strAction = Request.QueryString("Action")
strUserID = Request.QueryString("UserID")
strDescription = Request.QueryString("Description")
strComments = Request.QueryString("Comments")
strTable = Request.QueryString("Table")
lngRecordID = Request.QueryString("RecordID")
If Len(lngRecordID) = 0 Then lngRecordID = -1
mstrDetails = Request.QueryString("Details")
mstrChangeDate = Request.QueryString("ChangeDate")

mstrDetails = Replace(mstrDetails, "[PDSGN]sq[PDSGN]", "'")
mstrDetails = Replace(mstrDetails, "[PDSGN]dq[PDSGN]", """")
mstrDetails = Replace(mstrDetails, "[PDSGN]ba[PDSGN]", "|")
mstrDetails = Replace(mstrDetails, "[PDSGN]ex[PDSGN]", "!")
mstrDetails = Replace(mstrDetails, "[PDSGN]", "#")

If strAction = "Write" Then
    Set gadoCmd = GetAdoCmd("spActivityAuditAdd")
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, strUserID
        AddParmIn gadoCmd, "@ActionDescription", adVarChar, 80, strDescription
        AddParmIn gadoCmd, "@ActionComments", adVarChar, 5000, strComments
        AddParmIn gadoCmd, "@TableName", adVarChar, 50, strTable
        AddParmIn gadoCmd, "@TableRecordID", adInteger, 0, lngRecordID
        'ShowCmdParms(gadoCmd) '***DEBUG
        gadoCmd.Execute
    Set gadoCmd = Nothing
ElseIf strAction = "Read" Then
    ' Get Audit Activity records
    Set adRs = Server.CreateObject("ADODB.Recordset")
    Set gadoCmd = GetAdoCmd("spActivityAuditList")
        AddParmIn gadoCmd, "@TableName", adVarChar, 50, strTable
        AddParmIn gadoCmd, "@TableRecordID", adInteger, 0, lngRecordID
        AddParmIn gadoCmd, "@UserLogin", adVarChar, 20, Null
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, Null
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, Null
        AddParmIn gadoCmd, "@AuditAction", adVarChar, 100, NULL
        'ShowCmdParms(gadoCmd) '***DEBUG
        adRs.Open gadoCmd, , adOpenForwardOnly, adLockReadOnly
    Set gadoCmd = Nothing

    adRs.Sort = "[Date Of Action] DESC"
    Do While Not adRs.EOF
        mstrComments = ConvertCRLFToBR(adRs("Changes").Value)
        mstrComments = Replace(mstrComments,"^","*")
        mstrComments = Replace(mstrComments,"|","[BAR.]")
        mstrAudit = mstrAudit & adRs("Audit Record ID").Value & "^" & _
            adRs("Date Of Action").Value & "^" & _
            adRs("User ID").Value & "^" & _
            adRs("Action").Value & "^" & _
            mstrComments & "^" & _
            adRs("audTableName").Value & "^" & _
            adRs("Record ID").Value & "|"
        adRs.MoveNext
    Loop
    adRs.Close
ElseIf strAction = "Details" Then
End If

Function ConvertCRLFToBR(strText)
    Dim strTemp
    Dim intI
    
    If IsNull(strText) Then
        ConvertCRLFToBR = ""
        Exit Function
    End If
    strTemp = ""
    For intI = 1 To Len(strText)
        If Asc(Mid(strText, intI, 1)) = 13 Then
            strTemp = strTemp & "[linebreak]"
        Else
            If Asc(Mid(strText, intI, 1)) <> 10 Then
                strTemp = strTemp & Mid(strText, intI, 1)
            End If
        End If
    Next
    ConvertCRLFToBR = strTemp
End Function

%>
<HTML><HEAD>
    <TITLE>Audit History</TITLE>
    <meta name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
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
    <!--#include file="IncTableStyles.asp"-->
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
    Dim strValue
    Dim intLast
    Dim intCnt
    Dim intDelim
    Dim strRecords
    Dim mdctAudit
    
    Select Case "<%=strAction%>"
        Case "Write"
            window.close
        Case "Details"
        Case "Read"
            Set mdctAudit = CreateObject("Scripting.Dictionary")
            strRecords ="<% = mstrAudit %>"
            If Len(strRecords) > 1 Then
	            intLast = 0
	            intCnt = -1
	            ' Load array of items from string value
	            Do While True
		            intDelim = Instr(intLast + 1,strRecords,"|")
		            strValue = Mid(strRecords, intLast + 1, intDelim - (intLast + 1))
		            intCnt = intCnt + 1
                    mdctAudit.Add Parse(strValue,"^",1),strValue
		            intLast = intDelim
		            If intLast = Len(strRecords) Then Exit Do
	            Loop
	        Else
	            mdctAudit.Add 0,"0^^^^^^"
	        End If
            window.returnvalue = mdctAudit
            window.close
    End Select
End Sub

Sub cmdClose_onclick()
    window.close
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub window_onbeforeprint()
    cmdClose.style.left = -1000
    cmdPrint.style.left = -1000
    divAuditHistory.style.borderStyle="none"
End Sub

Sub window_onafterprint()
    cmdClose.style.left = 500
    cmdPrint.style.left = 390
    divAuditHistory.style.borderStyle="solid"
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:white;" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<% If strAction = "Details" Then %>
    <SPAN id=lblHeading class=DefLabel
        style="FONT-SIZE:10pt; HEIGHT:20; WIDTH:590; TOP:3; LEFT:10; TEXT-ALIGN:left;position:absolute">
        Audit History for change made on <B><%=mstrChangeDate%></B> by <B><%=strUserID%></B> for review ID <B><%=lngRecordID%></B>
    </SPAN>
    <DIV id=divAuditHistory Class=TableDivArea style="LEFT:10; TOP:25; WIDTH:590; HEIGHT:355; 
        OVERFLOW:auto; FONT-WEIGHT:normal; BACKGROUND-COLOR:<%=gstrBackColor%>;z-index:1">
        <TABLE id=tblAudit Border=0 Rules=rows Width=570 CellSpacing=0 
            Style="position:absolute;overflow: hidden; TOP:0;">
            <THEAD id=tbhAudit style="height:17">
                <TR id=thrAudit>
                    <TD class=CellLabel id=thcAuditC0 style="">Entry Name</TD>
                    <TD class=CellLabel id=thcAuditC1 style="">Value Before Change</TD>
                    <TD class=CellLabel id=thcAuditC2 style="">Value After Change</TD>
                </TR>
            </THEAD>
            <TBODY id=tbdAudit>
        <%
            Dim intI
            Dim strRecord
            For intI = 1 To 1000
                If Parse(mstrDetails,"[BAR.]",intI) = "" Then Exit For
                
                strRecord = Parse(mstrDetails,"[BAR.]",intI)
                
                Response.Write "<TR id=tdrAudit" & intI & ">"
                Response.Write "    <TD class=TableDetail id=tdcAuditC0" & intI & " style="""">" & Parse(strRecord,"*",1) & "</TD>" 
                Response.Write "    <TD class=TableDetail id=tdcAuditC1" & intI & " style="""">" & Parse(strRecord,"*",2) & "</TD>"
                Response.Write "    <TD class=TableDetail id=tdcAuditC2" & intI & " style="""">" & Parse(strRecord,"*",3) & "</TD>"
                Response.Write "</TR>"
            Next  
        %>
            </TBODY>
        </TABLE>
    </DIV>
    <INPUT type="button" value="Close" ID=cmdClose NAME=cmdClose
        title="Close this window" tabindex=-1
        style="position:absolute;LEFT:500; WIDTH:100; TOP:390">

    <INPUT type="button" value="Print" ID=cmdPrint NAME="cmdPrint"
        title="Print" tabindex=-1
        style="position:absolute;LEFT:390; WIDTH:100; TOP:390">
<%End If%>
</BODY>
</HTML>
<!--#include file="IncCmnCliFunctions.asp"-->
