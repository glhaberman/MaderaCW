<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<!--#include file="IncCnn.asp"-->
<%
Dim adCmd, oFso, oImport
Dim adRs, adRsAuthTypes
Dim strRecord, strType, strShortName, strLongName
Dim intStart, intI, intLen, intID

Set oFso = CreateObject("Scripting.FileSystemObject")
Set oImport = oFso.OpenTextfile("C:\inetpub\wwwroot\MaderaCW\CausalFactorDescriptionsBackFromMadera2.csv", 1)
' Read from the file and display the results.
Do While oImport.AtEndOfStream <> True
    strRecord = oImport.ReadLine
    intID = Parse(strRecord,",",1)
    If intID <> "" Then
        intStart = InStr(1,strRecord,",")
        strShortName = Mid(strRecord, intStart+1,Len(strRecord)-intStart)
        strShortName = Replace(strShortName,"""","")
        strShortName = Replace(strShortName,"'","`")
        strShortName = Trim(strShortName)

        response.Write "intID=" & intID & "<BR>" '& " --- " & strShortName & "<BR>"

        Set gadoCmd = GetAdoCmd("spTEMPUpdateFactorDesc")
            AddParmIn gadoCmd, "@FactorID", adInteger, 0, intID
            AddParmIn gadoCmd, "@Description", adVarchar, 5000, strShortName
            'ShowCmdParms(gadoCmd) '***DEBUG
            gadoCmd.Execute
        Set gadoCmd = Nothing
    End If
Loop
Response.Write "DONE <BR>"

Function CleanText(strText)
    'strText = Replace(strText, "^", Chr(9) & "#ca#")
    'strText = Replace(strText, "|", Chr(9) & "#ba#")
    'strText = Replace(strText, "*", Chr(9) & "#as#")
    'strText = Replace(strText, "!", Chr(9) & "#ex#")
    strText = Replace(strText, "'", Chr(9) & "#sq#")
    strText = Replace(strText, """", Chr(9) & "#dq#")
    strText = Replace(strText, Chr(13), "[l`b]")
    strText = Replace(strText, Chr(10), "")

    CleanText = strText
End Function

%>
<HTML><HEAD>
    <TITLE>Import CHOICES Authorities</TITLE>
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

Sub window_onload()
    
End Sub

</SCRIPT>
<BODY id=PageBody style="BACKGROUND-COLOR:white;" bottomMargin=10 leftMargin=10 topMargin=10 rightMargin=10>
</BODY>
</HTML>
<!--#include file="IncCmnCliFunctions.asp"-->
