<%@ LANGUAGE="VBScript" EnableSessionState=False%>
<%Option Explicit%>
<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'     Name: EmployeeAddEdit.asp                                              '
'  Purpose: The data entry screen for maintaining the appliction's staff    '
'           (worker) table.                                                 '
'           This form is only available to admin users.                     '
' Includes:                                                                 '
'   IncCnn.asp          - Connects to the database.                         '
'   IncValidUser.asp    - Code to lookup user ID and PW sent to the page.   '
'   IncDefStyles.asp    - Contains DHTML styles common in the application.  '
'==========================================================================='
%>
<!--#include file="IncCnn.asp"-->
<%
Dim strSQL
Dim mstrPageTitle
Dim adRs
Dim adCmd
Dim mstrHelpText

mstrPageTitle = "Case Review Online Help"

Set adRs = Server.CreateObject("ADODB.Recordset")
Set adCmd = Server.CreateObject("ADODB.Command")

mstrHelpText = ""
mstrHelpText = mstrHelpText & "<center><b>Employee Add/Edit Screen</b></center><br>"
mstrHelpText = mstrHelpText & "The Employee Add/Edit screen is where new employees are added to the Case Review system, or the information for an existing employee is modified.  It is also the screen where an employee may be placed into or removed from a position in the organization.<br><br>"
mstrHelpText = mstrHelpText & "<b>Record ID:</b><br>"
mstrHelpText = mstrHelpText & "The Record ID is an internal identifier assigned to the employee record at the time it was created.  It is displayed for informational purposes only.<br><br>"
mstrHelpText = mstrHelpText & "<b>Employee ID:</b><br>"
mstrHelpText = mstrHelpText & "Enter the personnel number for the employee.  This is the unique number used by your organization to identify a person.<br><br>" 
mstrHelpText = mstrHelpText & "<b>Employee Name:</b><br>"
mstrHelpText = mstrHelpText & "Enter the employee's last name, first name, and middle name or middle initial.<br>"
mstrHelpText = mstrHelpText & "&nbsp&nbsp&nbsp Maximum number of characters for [Last Name]: 50<br>"
mstrHelpText = mstrHelpText & "&nbsp&nbsp&nbsp Maximum number of characters for [First Name]: 25<br>"
mstrHelpText = mstrHelpText & "&nbsp&nbsp&nbsp Maximum number of characters for [Middle Name]: 25<br><br>"
mstrHelpText = mstrHelpText & "<b>Employment Dates:</b><br>"
mstrHelpText = mstrHelpText & "The Employment Dates section of the screen is used to enter the starting and ending dates for the employee's time with the organization.<br>"
mstrHelpText = mstrHelpText & "When adding a new employee enter the [Starting Date] only, leaving the [Ending Date] blank.<br>"
mstrHelpText = mstrHelpText & "When an employee leaves the organization and will no longer be a part of the Case Review system, edit the employee's record and fill in the [Ending Date].  An employee who has been ended will not appear in dropdown lists on new reviews, but their history will remain a part of the system.<br><br>"
mstrHelpText = mstrHelpText & "<b>Position Information:</b><br>"
mstrHelpText = mstrHelpText & "A Position in the case review system corresponds to a job or full-time position in the real world organization.  Different states and counties may use different terms - job, caseload, desk, worker code, etc., but the essential point is that the position record represents the workload begin done, while the employee record represents the person doing the work.<br><br>"
mstrHelpText = mstrHelpText & "The Position Information section consists of a dropdown list with all positions in the organization that are currently vacant (not filled by a current employee), the name of the supervisor for the position currently selected in the dropdown list, and a field to enter the date on which the employee started working in the selected position.<br><br>"
mstrHelpText = mstrHelpText & "By selecting a position in the dropdown list, the employee is placed in that position - in other words, the employee's name is associated with that particular job.  The value entered in the [Starting Date In Position] field tells the system when the employee began working in that job.<br><br>"
mstrHelpText = mstrHelpText & "When editing an existing employee record, the position that was selected on entering the screen is displayed with an asterisk.<br>"

mstrHelpText = mstrHelpText & "<br>"

Response.ExpiresAbsolute = Now - 5
%>
<HTML>
<HEAD>
    <meta name=vs_targetSchema content="HTML 4.0">
    <TITLE><%=Trim(gstrOrgAbbr & " " & gstrAppName) & " Online Help"%></TITLE>
    <!--#include file="IncDefStyles.asp"-->
</HEAD>

<SCRIPT ID=clientEventHandlersVBS LANGUAGE=vbscript>
Sub window_onload()
    'Fill In
End Sub

Sub cmdPrint_onclick()
    window.print
End Sub

Sub cmdClose_onclick()
    window.close    
End Sub
</SCRIPT>

<BODY id=PageBody 
    bottomMargin=5 
    leftMargin=5
    topMargin=5 
    rightMargin=5>
    
    <DIV id=Header
        class=DefTitleArea
        style="WIDTH:520; height:40">

        <SPAN id=lblAppTitleHiLight
            class=DefTitleTextHiLight
            style="WIDTH:520"><%=mstrPageTitle%>
        </SPAN>
        
        <SPAN id=lblAppTitle
            class=DefTitleText
            style="WIDTH:520"><%=mstrPageTitle%>
        </SPAN>
    </DIV>

    <DIV id=PageFrame
        class=DefPageFrame
        style="HEIGHT:320; WIDTH:520; TOP:51">

        <DIV id=txtHelpText
            style="LEFT:10; HEIGHT:265; WIDTH:500; TOP:10;
                BACKGROUND-COLOR:white; 
                OVERFLOW:scroll; 
                padding:10;
                padding-top:5" 
            tabIndex=1><%=mstrHelpText%>
        </DIV>

        <BUTTON id=cmdPrint
            class=DefBUTTON
            title="Print help" 
            style="LEFT:360; TOP:290; HEIGHT: 20; WIDTH:70"
            accessKey=P
            tabIndex=3><u>P</u>rint
        </BUTTON>
        <BUTTON id=cmdClose
            class=DefBUTTON
            title="Close the form" 
            style="LEFT:440; TOP:290; HEIGHT: 20; WIDTH:70"
            tabIndex=4>Close
        </BUTTON>
    </DIV>

</BODY>
</FORM>
</HTML>
<%
'adRs.Close
Set adRs = Nothing
gadoCon.Close
Set gadoCon = Nothing
%>