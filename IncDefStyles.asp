<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncDefStyles.asp                                                 '
' Purpose: This include file contains common styles.                        '
'                                                                           '
'==========================================================================='
%>

    <STYLE id=DefaultStyles type="text/css" rel="stylesheet">
        BODY
            {
            margin:1;
            position: absolute; 
            FONT-SIZE: 10pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            OVERFLOW: auto; 
            BACKGROUND-COLOR: <%=gstrPageColor%>
            }
            
        DIV {position: absolute}
        
        SPAN {position: absolute}

        TEXTAREA
            {
            padding-left: 2;
            position: absolute; 
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            OVERFLOW: hidden; 
            TEXT-ALIGN: left; 
            HEIGHT: 18px; 
            BACKGROUND-COLOR: white
            }

        SELECT
            {
            position: absolute; 
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            HEIGHT: 18px; 
            BACKGROUND-COLOR: white
            }

        SELECT
            {
            position: absolute; 
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            HEIGHT: 18px; 
            BACKGROUND-COLOR: white
            }

        INPUT
            {
            position: absolute; 
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            HEIGHT: 18px;
            TEXT-ALIGN:center
            }


        BUTTON
            {
            position: absolute; 
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>
            }
        
        .DefLabel
            {
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>; 
            PADDING-LEFT: 1px; 
            TEXT-ALIGN: left; 
            OVERFLOW: visible;
            HEIGHT: 15;
            CURSOR: default
            }
            
        .DefBUTTON
            {
            POSITION: absolute;
            FONT-SIZE: 8pt; 
            FONT-FAMILY: <%=gstrTextFont%>;
            HEIGHT: 25;
            WIDTH: 155;
            <%If gstrDefButtonColor <> vbNullString Then %>
                COLOR: <%=gstrDefButtonText%>;
                BACKGROUND-COLOR: <%=gstrDefButtonColor%>
            <%End If%>
            }
        
        .DefRectangle
            {
            POSITION: absolute;
            BACKGROUND-COLOR: <%=gstrPageColor%>;
            BORDER-COLOR: <%=gstrBorderColor%>;
            BORDER-STYLE: solid;
            BORDER-WIDTH:1;
            OVERFLOW: hidden
            }
        
        .DefTitleArea
            {
            position: absolute;
            BACKGROUND-COLOR: <%=gstrBackColor%>;
            BORDER-STYLE: solid;
            BORDER-WIDTH: 1px;
            BORDER-COLOR: <%=gstrBorderColor%>;
            COLOR: <%=gstrForeColor%>;
            TOP: 10;
            HEIGHT: 40;
            LEFT: 10;
            }

        .DefTitleTextHiLight
            {
            position: absolute;
            TOP: 3;
            LEFT: 1;
            COLOR: <%=gstrAccentColor%>;
            FONT-FAMILY: <%=gstrTitleFont%>;
            FONT-SIZE: <%=gstrTitleFontSize%>;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold
            }

        .DefTitleText
            {
            position: absolute;
            TOP: 2;
            LEFT: 0;
            COLOR: <%=gstrTitleColor%>;
            FONT-FAMILY: <%=gstrTitleFont%>;
            FONT-SIZE: <%=gstrTitleFontSize%>;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold
            }
            
        .DefPageFrame
            {
            BORDER-STYLE: solid;
            BORDER-WIDTH: 1px;
            BORDER-COLOR: <%=gstrBorderColor%>;
            COLOR: <%=gstrForeColor%>;
            BACKGROUND-COLOR: <%=gstrBackColor%>;
            LEFT: 10
            }
    </STYLE>