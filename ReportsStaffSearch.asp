<%@ LANGUAGE="VBScript" EnableSessionState=False%><%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: PosEmpOpenPositions.asp                                         '
'  Purpose: This screen allows the user to view all open position records   '
'  in a tree view format.                                                   '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim adRs
Dim mstrPageTitle
Dim mintTblWidth
Dim mlngAliasID
Dim mblnUserAdmin
Dim mblnUserQA
Dim mstrWorkerList
Dim mstrAuthByList
Dim mstrType
Dim mstrStaffName
Dim mintMngLvlID
Dim mdctSups
Dim mstrStaffID
Dim mstrStaffInfo
Dim mstrAllowAll    'Adds All as the first option in table
Dim mstrBackColor   'Sets the background color of the table to allow it to match parent page
Dim mintWidth       'Sets width of div/table to allow it to match parent page controls
Dim mblnCTR         'Indicates whether or not the logged in user has the CTR role and/or is in CTR Mng Level
Dim mstrMngLvlIDs
Dim mstrUserID, mstrManager, mstrDirector, mstrOffice, mstrSupervisor
Dim mdtmStartDate, mdtmEndDate

mstrStaffID = Request.QueryString("StaffID")
mlngAliasID = Request.QueryString("AliasID")
mblnUserAdmin = Request.QueryString("UserAdmin")
mstrUserID = Request.QueryString("UserID")
mblnUserQA = Request.QueryString("UserQA")
mstrMngLvlIDs = Request.QueryString("MngLvlIDs")
mstrStaffName = Request.QueryString("StaffName")
mstrAllowAll = Request.QueryString("AllowAll")
If Len(mstrAllowAll) = 0 Then mstrAllowAll = "N"
mstrBackColor = Request.QueryString("BackColor")
If Len(mstrBackColor) = 0 Then mstrBackColor = "lightyellow"
mintWidth = Request.QueryString("Width")
If Len(mintWidth) = 0 Then mintWidth = 194
mstrType = Request.QueryString("Type")
mstrManager = Request.QueryString("Manager")
mstrOffice = Request.QueryString("Office")
mstrDirector = Request.QueryString("Director")
mstrSupervisor = Request.QueryString("Supervisor")
mdtmStartDate = Request.QueryString("StartDate")
If Len(mdtmStartDate) = 0 Then mdtmStartDate = Null
mdtmEndDate = Request.QueryString("EndDate")
If Len(mdtmEndDate) = 0 Then mdtmEndDate = Null
Select Case mstrType
    Case "txtWorker"
        mintMngLvlID = 121
    Case "txtArcWorker"
        mintMngLvlID = 1211
    Case "txtSupervisor"
        mintMngLvlID = 124
    Case "txtArcSupervisor"
        mintMngLvlID = 1241
    Case "txtReviewer"
        mintMngLvlID = 127
    Case "txtReReviewer"
        mintMngLvlID = 1278
    Case "txtEmployee"
        mintMngLvlID = 1210
End Select
mstrPageTitle = "Select " & mstrType

Set mdctSups = CreateObject("Scripting.Dictionary")
%>
<HTML>
<HEAD>
    <META name=vs_targetSchema content="http://schemas.microsoft.com/intellisense/ie5">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName)%></TITLE>
	<!--#include file="IncDefStyles.asp"-->
    <!--#include file="IncTableStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Dim mintSelectedRow
Dim mintColoredRow

Sub window_onload()
    mintSelectedRow = -1
    mintColoredRow = -1
    divPositions.focus
    divPositions.style.left = 0
    divLoading.style.left = -1000
    
    If "<% = mstrAllowAll %>" = "Y" Then
        tblStaff.rows("tbrStaff1").cells(0).innerText = "<All>"
    End If
End Sub

Sub Result_onmouseover(intRow)
    Dim objCell
    
    For Each objCell In tblStaff.rows("tbrStaff" & intRow).cells
        objCell.style.backgroundcolor = "darkolivegreen"
        objCell.style.color = "white"
    Next
    
    If mintColoredRow > 0 And mintColoredRow <> intRow Then
        For Each objCell In tblStaff.rows("tbrStaff" & mintColoredRow).cells
            objCell.style.backgroundcolor = "transparent"
            objCell.style.color = "black"
        Next
    End If
    mintColoredRow = intRow
End Sub

Sub Result_onclick(intRow)
    Dim objCell
    
    For Each objCell In tblStaff.rows("tbrStaff" & intRow).cells
        objCell.style.backgroundcolor = "darkolivegreen"
        objCell.style.color = "white"
    Next
    
    If mintSelectedRow > 0 Then
        For Each objCell In tblStaff.rows("tbrStaff" & mintSelectedRow).cells
            objCell.style.backgroundcolor = "transparent"
            objCell.style.color = "black"
        Next
    End If
    mintSelectedRow = intRow
    
    Call SelectAndClose(mintSelectedRow)
End Sub

Sub LostFocus()
    If mintSelectedRow > 0 Then
        Call SelectAndClose(mintSelectedRow)
    Else
        Call cmdCancel_onclick()
    End If
End Sub

Sub SelectAndClose(intRow)
    Dim strEmpID, strReturn
    
    strReturn = tblStaff.rows("tbrStaff" & mintSelectedRow).innerText & "^" & document.all("txtStaffDates" & mintSelectedRow).value
    Call window.parent.StaffLookUpClose(strReturn)
    divPositions.style.left = -1000
    divLoading.style.left = 0
End Sub

Sub cmdSelect_onclick()
    If mintSelectedRow <= 0 And mintColoredRow > 0 Then mintSelectedRow = mintColoredRow
    If mintSelectedRow > 0 Then Call SelectAndClose(mintSelectedRow)
End Sub

Sub cmdCancel_onclick()
    Call window.parent.StaffLookUpClose("[CANCEL]")
    divPositions.style.left = -1000
    divLoading.style.left = 0
End Sub

Sub divPositions_onkeydown()
    Dim intTop
    <%
    'This code controls the behavior in the results DIV when the Up arrow,
    'Down arrow, Home, and End keys are pressed.  This code changes the
    'selected item as the user moves up and down in the list:
    %>
    If InStr(tblStaff.Rows(0).innerHTML,"**no staff found**") > 0 Then Exit Sub
    intTop = -1    
    If IsNumeric(mintColoredRow) Then
        Select Case Window.Event.keyCode
            Case 36 'home
                Window.event.returnValue = False
                tblStaff.rows(1).scrollIntoView
                Call MoveColoredRow(1)
                intTop = 1
            Case 35 'end
				Window.event.returnValue = False
                Call MoveColoredRow(tblStaff.Rows.Length)
                intTop = tblStaff.Rows.Length
            Case 38 'Up
                If mintColoredRow > 1 Then
					Window.event.returnValue = False
                    Call MoveColoredRow(mintColoredRow - 1)
                    intTop = mintColoredRow - 1 
                End If
            Case 40 'Down
                If Cint(mintColoredRow) < Cint(tblStaff.Rows.Length) Then
					Window.event.returnValue = False
                    Call MoveColoredRow(mintColoredRow + 1)
                    intTop = mintColoredRow + 1 
                End If
            Case 13
                Call cmdSelect_onclick()
            Case 27
                Call cmdCancel_onclick()
        End Select
        If intTop >= 0 Then
            txtFocus.style.top = intTop * 15
            txtFocus.focus
        End If
    End If
End Sub
Sub MoveColoredRow(intRow)
    Dim objCell
    
    If intRow <= 0 Then intRow = 1
    For Each objCell In tblStaff.rows("tbrStaff" & intRow).cells
        objCell.style.backgroundcolor = "darkolivegreen"
        objCell.style.color = "white"
    Next
    
    If mintColoredRow > 0 Then
        For Each objCell In tblStaff.rows("tbrStaff" & mintColoredRow).cells
            objCell.style.backgroundcolor = "transparent"
            objCell.style.color = "black"
        Next
    End If
    mintColoredRow = intRow
End Sub


</SCRIPT>
<!--#include file="IncCmnCliFunctions.asp"-->

<BODY id=BODY1>
    <DIV id=divLoading 
        style="BACKGROUND-COLOR:<% = mstrBackColor %>;LEFT:-1000; WIDTH:<% = mintWidth %>; TOP:0; HEIGHT:144;">
        <BR><BR>Searching...
    </DIV>
	<DIV id=divPositions class=TableDivArea
        style="LEFT:0; WIDTH:<% = mintWidth %>; TOP:0; HEIGHT:144;background-color:<% = mstrBackColor %>" tabIndex=5 >
        <%
        Dim strStaffName
        Dim lngStaffPosID
        Dim strOptions
        Dim intI
        Dim strDisabled
        Dim strHiddenIDs
        
        Select Case mintMngLvlID
            Case 1211, 1241
                If mstrDirector = 0 Then mstrDirector = Null
                If mstrOffice = 0 Then mstrOffice = Null
                If mstrManager = 0 Then mstrManager = Null
                If Len(mstrSupervisor) = 0 Then mstrSupervisor = Null
                Set gadoCmd = GetAdoCmd("spArchiveGetStaffingWS")
                If mintMngLvlID = 1211 Then
                    AddParmIn gadoCmd, "@RoleName", adVarchar, 100, "Worker"
                Else
                    AddParmIn gadoCmd, "@RoleName", adVarchar, 100, "Supervisor"
                End If
                AddParmIn gadoCmd, "@Director", adVarChar, 50, mstrDirector
                AddParmIn gadoCmd, "@Office", adVarChar, 50, mstrOffice
                AddParmIn gadoCmd, "@Manager", adVarChar, 50, mstrManager
                AddParmIn gadoCmd, "@Supervisor", adVarChar, 150, mstrSupervisor
                AddParmIn gadoCmd, "@StaffName", adVarChar, 50, mstrStaffName
                AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, mdtmStartDate
                AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, mdtmEndDate
           Case Else
                Set gadoCmd = GetAdoCmd("spGetOptStaffList")
                gadoCmd.CommandTimeout = 180
                
                AddParmIn gadoCmd, "@AliasID", adInteger, 0, mlngAliasID
                AddParmIn gadoCmd, "@UserID", adVarChar, 20, mstrUserID
                AddParmIn gadoCmd, "@Admin", adBoolean, 0, mblnUserAdmin
                AddParmIn gadoCmd, "@QA", adBoolean, 0, mblnUserQA
                AddParmIn gadoCmd, "@MngLvlID", adInteger, 0, mintMngLvlID
                AddParmIn gadoCmd, "@StaffName", adVarChar, 50, mstrStaffName
                AddParmIn gadoCmd, "@StaffID", adVarChar, 20, Null 'mstrStaffID
                AddParmIn gadoCmd, "@Supervisor", adVarChar, 150, mstrSupervisor
                AddParmIn gadoCmd, "@Manager", adVarChar, 50, mstrManager
                AddParmIn gadoCmd, "@Office", adVarChar, 50, mstrOffice
                AddParmIn gadoCmd, "@Director", adVarChar, 50, mstrDirector
                AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, mdtmStartDate
                AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, mdtmEndDate
        End Select        
        'Call ShowCmdParms(gadoCmd) '***DEBUG
        Set adRs = GetAdoRs(gadoCmd)

        adRs.Sort = "StaffName"
        mintTblWidth = mintWidth - 20
        Response.Write "<Table ID=tblStaff Border=0 Rules=rows Cols=2 Width=" & mintTblWidth & " CellSpacing=0 Style=""overflow: hidden; TOP:0;background-color:transparent""> " & vbCrLf
        Response.Write "<TBODY ID=tbdStaff> " & vbCrLf
        intI = 0
        strDisabled = ""
        strHiddenIDs = ""
        If mstrAllowAll = "Y" Then
		    intI = intI + 1
            strStaffName = "All"
		    Response.Write "    <TR ID=tbrStaff" & intI & " class=TableRow onclick=Result_onclick(" & intI & ")> " & vbCrLf
		    Response.Write "        <TD ID=tbcStaffC1" & intI & " class=TableDetail style=""background-color:transparent;border-style:none;cursor:hand"" >" & strStaffName & "</TD>" & vbCrLf
		    Response.Write "    </TR>" & vbCrLf
            strOptions = strOptions & "<OPTION VALUE=1>All</OPTION>"
        End If
	    Do While Not adRs.EOF
		    intI = intI + 1
		    lngStaffPosID = intI
		    strStaffName = adRs.Fields("StaffName").value
			
		    Response.Write "    <TR ID=tbrStaff" & lngStaffPosID & " class=TableRow onmouseover=Result_onmouseover(" & intI & ") onclick=Result_onclick(" & lngStaffPosID & ") > " & vbCrLf
		    Response.Write "        <TD ID=tbcStaffC1" & lngStaffPosID & " class=TableDetail style=""background-color:transparent;border-style:none;cursor:hand"" >" & strStaffName & "</TD>" & vbCrLf
		    Response.Write "    </TR>" & vbCrLf
		    strHiddenIDs = strHiddenIDs & "<INPUT type=""hidden"" id=txtStaffDates" & lngStaffPosID & " VALUE=""" & adRs.Fields("StartDate").value & "^" & adRs.Fields("EndDate").value & """>" & vbCrLf
		    adRs.MoveNext 
	    Loop
		
	    If intI = 0 Then 
	        strDisabled = "disabled"
            lngStaffPosID = 1
            strStaffName = "no matches [Close]"
            strOptions = "<OPTION VALUE=" & lngStaffPosID & ">^</OPTION>"
	        Response.Write "    <TR ID=tbrStaff" & lngStaffPosID & " class=TableRow onclick=Result_onclick(" & lngStaffPosID & ") > " & vbCrLf
	        Response.Write "        <TD ID=tbcStaffC1" & lngStaffPosID & " class=TableDetail style=""background-color:transparent"" >" & strStaffName & "</TD>" & vbCrLf
	        Response.Write "    </TR>" & vbCrLf
		    strHiddenIDs = strHiddenIDs & "<INPUT type=""hidden"" id=txtStaffDates" & lngStaffPosID & " VALUE=""^"">" & vbCrLf
        End if
        Response.Write "</TBODY> </TABLE>" & vbCrLf & vbCrLf
        Response.Write strHiddenIDs
        %>
        <INPUT type="text" ID=Text1 NAME="txtFocus" style="left:0;width:1;height:1;z-index:-1">
    </DIV>
</BODY>

<FORM NAME="Form" METHOD="Post" STYLE="VISIBILITY: hidden" ID=Form1>
    <%
	WriteFormField "StaffInfo", mstrStaffInfo
    %>
</FORM>
</HTML>
<!--#include file="IncCmnFormFields.asp"-->
<!--#include file="IncWriteFormField.asp"-->
