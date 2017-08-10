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
Dim mdtmStartDate
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
Dim mstrUserID

mstrStaffID = Request.QueryString("StaffID")
mlngAliasID = Request.QueryString("AliasID")
mblnUserAdmin = Request.QueryString("UserAdmin")
mstrUserID = Request.QueryString("UserID")
mblnUserQA = Request.QueryString("UserQA")
mstrMngLvlIDs = Request.QueryString("MngLvlIDs")
mdtmStartDate = Request.QueryString("StartDate")
mstrStaffName = Request.QueryString("StaffName")
mstrAllowAll = Request.QueryString("AllowAll")
If Len(mstrAllowAll) = 0 Then mstrAllowAll = "N"
mstrBackColor = Request.QueryString("BackColor")
If Len(mstrBackColor) = 0 Then mstrBackColor = "lightyellow"
mintWidth = Request.QueryString("Width")
If Len(mintWidth) = 0 Then mintWidth = 194
mstrType = Request.QueryString("Type")
Select Case mstrType
    Case "txtManager"
        mintMngLvlID = 125
    Case "txtWorker"
        mintMngLvlID = 121
    Case "txtSupervisor"
        mintMngLvlID = 124
    Case "txtReviewer"
        mintMngLvlID = 127
    Case "txtReviewerList"
        mintMngLvlID = 1279
    Case "txtReReviewerList"
        mintMngLvlID = 1278
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
    
    strReturn = tblStaff.rows("tbrStaff" & mintSelectedRow).innerText
    'Select Case "<%=mintMngLvlID%>"
    '    Case "121", "128"
    '        strReturn = strReturn & " -- " & document.all("txtStaffEmpID" & mintSelectedRow).value
    'End Select
    Call window.parent.StaffLookUpClose(strReturn)
    divPositions.style.left = -1000
    divLoading.style.left = 0
End Sub

Sub cmdSelect_onclick()
    If mintSelectedRow <= 0 And mintColoredRow > 0 Then mintSelectedRow = mintColoredRow
    If mintSelectedRow > 0 Then Call SelectAndClose(mintSelectedRow)
End Sub

Sub cmdCancel_onclick()
    Call window.parent.StaffLookUpClose("")
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
        
        Set gadoCmd = GetAdoCmd("spGetOptStaffList")
        gadoCmd.CommandTimeout = 180
        
        AddParmIn gadoCmd, "@AliasID", adInteger, 0, mlngAliasID
        AddParmIn gadoCmd, "@UserID", adVarChar, 20, mstrUserID
        AddParmIn gadoCmd, "@Admin", adBoolean, 0, mblnUserAdmin
        AddParmIn gadoCmd, "@QA", adBoolean, 0, mblnUserQA
        AddParmIn gadoCmd, "@MngLvlID", adInteger, 0, mintMngLvlID
        AddParmIn gadoCmd, "@StaffName", adVarChar, 50, mstrStaffName
        AddParmIn gadoCmd, "@StaffID", adVarChar, 20, mstrStaffID
        AddParmIn gadoCmd, "@Supervisor", adVarChar, 150, Null
        AddParmIn gadoCmd, "@Manager", adVarChar, 50, Null
        AddParmIn gadoCmd, "@Office", adVarChar, 50, Null
        AddParmIn gadoCmd, "@Director", adVarChar, 50, Null
        AddParmIn gadoCmd, "@StartDate", adDBTimeStamp, 0, Null
        AddParmIn gadoCmd, "@EndDate", adDBTimeStamp, 0, Null
        Set adRs = GetAdoRs(gadoCmd)

        mintTblWidth = mintWidth - 20
        strHiddenIDs = ""
        Response.Write "<Table ID=tblStaff Border=0 Rules=rows Cols=2 Width=" & mintTblWidth & " CellSpacing=0 Style=""overflow: hidden; TOP:0;background-color:transparent""> " & vbCrLf
        Response.Write "<TBODY ID=tbdStaff> " & vbCrLf
        intI = 0
        strDisabled = ""
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
		    If mintMngLvlID = 121 Or mintMngLvlID=128 Then
		        strHiddenIDs = strHiddenIDs & "<INPUT type=""hidden"" id=txtStaffEmpID" & lngStaffPosID & " VALUE=""" & adRs.Fields("StaffEmpID").value & """>" & vbCrLf
		    End If
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
        End if
        Response.Write "</TBODY> </TABLE>"
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
