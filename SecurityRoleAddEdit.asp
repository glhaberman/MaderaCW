<%
Dim mstrOriginalText
Dim mstrOriginalCode
Dim madoRs
Dim madCmd


%>

        <SPAN class=DefLabel 
            id=lblInstructions
            style="POSITION: absolute;
                FONT-WEIGHT: bold;
                LEFT: 500; 
                WIDTH: 280; 
                TOP: 5">Instructions:
        </SPAN>
        <SPAN class=DefLabel 
            id=lblInstructionsText
            style="POSITION: absolute;
                LEFT: 500; 
                WIDTH: 200;
                HEIGHT: 75; 
                OVERFLOW: hidden;
                TOP: 20">Move items in or out of the list of classifications on the right. Click [Save] to keep your changes. Click [Close] to abandon changes.
        </SPAN>

        <SPAN class=DefLabel 
            id=lblJobClasses
            style="POSITION: absolute;
                LEFT: 15; 
                WIDTH: 175; 
                TOP: 5">Security Roles:
        </SPAN>
        <SELECT id=lstSecurityRoles
            title="List of Available Security Roles"
            style="Z-INDEX: 2; 
                LEFT: 15; 
                WIDTH: 175; 
                POSITION: absolute; 
                TOP: 25; 
                HEIGHT: 215" 
                tabIndex=1 
                size=13 
                TYPE="select-one">
                <%Set madCmd = Server.CreateObject("ADODB.Command")
                With madCmd
                    .ActiveConnection = gadoCon
                    .CommandType = adCmdStoredProc
                    .CommandText = "spGetAllSecurityRoles"
                    If request.Form("logRecordID") = "" Then
						.Parameters.Append .CreateParameter("@logRecordID", adInteger, adParamInput, 0, NULL)
					Else
						.Parameters.Append .CreateParameter("@logRecordID", adInteger, adParamInput, 0, request.Form("logRecordID"))
					End IF 
                End With
                'Open a recordset from the query:
                Set madoRs = Server.CreateObject("ADODB.Recordset") 
                Call madoRs.Open(madCmd, , adOpenForwardOnly, adLockReadOnly)
                Do While Not madoRs.EOF And Not madoRs.BOF
                    Response.Write "<OPTION VALUE=" & madoRs.Fields("rolRoleGroup").Value & ">" & madoRs.Fields("rolRoleName").Value
                    madoRs.MoveNext
                Loop
                madoRs.Close
                Set madCmd = Nothing%>
        </SELECT>
        <SELECT id=lstSecurityRolesAtLoad
            title=""
            style="Z-INDEX: 2; visibility: hidden;
                LEFT: 15; 
                WIDTH: 175; 
                POSITION: absolute; 
                TOP: 80; 
                HEIGHT: 215" 
                tabIndex=1 
                size=13 
                TYPE="select-one">
        </SELECT>

        <SPAN class=DefLabel 
            id=lblAssignedClasses
            style="POSITION: absolute;
                LEFT: 305; 
                WIDTH: 175; 
                TOP: 5">Assigned Security Roles:
        </SPAN>
        <SELECT id=lstAssignedRoles
            title="List Current Security Roles"
            style="Z-INDEX: 2; 
                LEFT: 305; 
                WIDTH: 175; 
                POSITION: absolute; 
                TOP: 25; 
                HEIGHT: 215" 
                tabIndex=1 
                size=13 
                TYPE="select-one">
                <%
				If request.Form("logRecordID") <> "" Then
					Set madCmd = Server.CreateObject("ADODB.Command")
					With madCmd
						.ActiveConnection = gadoCon
						.CommandType = adCmdStoredProc
						.CommandText = "spGetSecurityRoles"
						.Parameters.Append .CreateParameter("@logRecordID", adInteger, adParamInput, 0, request.Form("logRecordID"))
					End With
					'Open a recordset from the query:
					Set madoRs = Server.CreateObject("ADODB.Recordset") 
					Call madoRs.Open(madCmd, , adOpenForwardOnly, adLockReadOnly)
					Do While Not madoRs.EOF And Not madoRs.BOF
						Response.Write "<OPTION VALUE=" & madoRs.Fields("rolRoleGroup").Value & ">" & madoRs.Fields("rolRoleName").Value
						madoRs.MoveNext
					Loop
					madoRs.Close
					Set madCmd = Nothing
				End IF
                %>
        </SELECT>
        <SELECT id=lstAssignedRolesAtLoad
            title=""
            style="Z-INDEX: 2; visibility: hidden;
                LEFT: 305; 
                WIDTH: 175; 
                POSITION: absolute; 
                TOP: 80; 
                HEIGHT: 215" 
                tabIndex=1 
                size=13 
                TYPE="select-one">
        </SELECT>

        <BUTTON class=DefBUTTON
            id=cmdLeftToRight title="Add new value to the list" 
            style="LEFT: 210;
                POSITION: absolute;
                TOP: 105;
                WIDTH: 70;
                HEIGHT: 20"
            accessKey=R
            tabIndex=7>--&gt
        </BUTTON>
        <BUTTON class=DefBUTTON
            id=cmdRightToLeft title="Delete the selected value" 
            style="LEFT: 210;
                POSITION: absolute;
                TOP: 135;
                WIDTH: 70;
                HEIGHT: 20"
            accessKey=R
            tabIndex=7>&lt--
        </BUTTON>
        
