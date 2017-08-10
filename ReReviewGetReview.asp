<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: ReReviewGetReview.asp                                           '
'  Purpose: This page is used fetch Review information for a re-review.    '
' Includes:                                                                 '
'   IncCnn.asp          - ADO database connection                           '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd
Dim adRs
Dim adRsElements, adRsComments
Dim lngReviewID
Dim intI
Dim mstrGlobalParms, mstrComment

lngReviewID = Request.QueryString("ReviewID")
mstrGlobalParms = Request.QueryString("GlobalParms")
If Len(lngReviewID) = 0 Then lngReviewID = 0

Set adCmd = GetAdoCmd("spReReviewGetReview")
Set adRsElements = Server.CreateObject("ADODB.Recordset")
Set adRsComments = Server.CreateObject("ADODB.Recordset")

adCmd.CommandTimeout = 180
    AddParmIn adCmd, "@ReviewID", adInteger, 0, lngReviewID
    AddParmIn adCmd, "@UserID", adVarChar, 20, Parse(mstrGlobalParms,"^",1)
    'Call ShowCmdParms(adCmd) '***DEBUG
Set adRs = GetAdoRs(adCmd)
If adRs.State = adStateOpen Then
    Set adRsElements = adRs.NextRecordset
    adRsElements.Filter = "ItemStatusID<>25"
    Set adRsComments = adRs.NextRecordset
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
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
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
</HEAD>

<SCRIPT LANGUAGE=vbscript>
Option Explicit

Sub window_onload
    Dim dctReturn
    Set dctReturn = CreateObject("Scripting.Dictionary")
    
<%
    If Not adRs.Eof Then
        Response.Write "dctReturn.Add ""ReviewMonth"", """ & adRs.Fields("rvwMonthYear").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""ReviewDate"", """ & adRs.Fields("rvwDateEntered").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""ReviewClass"", """ & adRs.Fields("ReviewClass").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""CaseName"", """ & adRs.Fields("ClientName").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""CaseNumber"", """ & adRs.Fields("rvwCaseNumber").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""ReviewStatus"", """ & adRs.Fields("ReviewStatus").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""ReviewerName"", """ & adRs.Fields("rvwReviewerName").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""WorkerName"", """ & adRs.Fields("rvwWorkerName").value & """" & vbCrLf
        'Response.Write "dctReturn.Add ""AuthorizedByName"", """ & adRs.Fields("rvwAuthByName").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""WorkerResponse"", """ & adRs.Fields("WorkerResponse").value & """" & vbCrLf
        Response.Write "dctReturn.Add ""ProgramsReviewed"", """ & adRs.Fields("ProgramsReviewed").value & """" & vbCrLf

        intI = 0        
        Do While Not adRsElements.Eof
            intI = intI + 1
            If adRsElements.Fields("rveTypeID").value = 2 Then
                adRsComments.Filter = "rvcScreenName='" & adRsElements.Fields("ElementName").value & "'"
                mstrComment = ""
                If adRsComments.RecordCount = 1 Then
                    mstrComment = ConvertCRLFToBR(adRsComments.Fields("rvcComments").value)
                End If
            Else
                mstrComment = ConvertCRLFToBR(adRsElements.Fields("rveComments").value)
            End If
            Response.Write "dctReturn.Add ""Element" & intI & """, """ & _
                adRsElements.Fields("ProgramName").value & "^" & _
                adRsElements.Fields("ElementName").value & "^" & _
                adRsElements.Fields("ItemStatus").value & "^" & _
                adRsElements.Fields("FactorName").value & "^" & _
                adRsElements.Fields("GroupID").value & "^" & _
                adRsElements.Fields("GroupName").value & "^" & _
                mstrComment & "^" & _
                adRsElements.Fields("rveProgramID").value & "^" & _
                adRsElements.Fields("rveElementID").value & "^" & _
                "" & "^" & _ 
                "" & "^" & _ 
                adRsElements.Fields("ReviewType").value & "^" & _ 
                adRsElements.Fields("rveTypeID").value & "^" & _ 
                adRsElements.Fields("rvfFactorID").value & """" & vbCrLf
            '2 Place holders above for ReReview Status and Comments
            adRsElements.MoveNext
        Loop
    Else
        Response.Write "dctReturn.Add ""NotFound"",""""" & vbCrLf
    End If
%>
    window.returnvalue = dctReturn
    window.close
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:#white; overflow:scroll" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
<BR><BR><BR>
<div id=divDisplay>
    Retrieving Review Information, please wait...
</div>

</BODY>
</HTML>
