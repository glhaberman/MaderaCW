<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncRptStyles.asp                                                 '
' Purpose: This include file contains report styles.                        '
'                                                                           '
'==========================================================================='
%>

    <STYLE id=DefaultStyles type="text/css" rel="stylesheet">
        .RptCriteriaCell
            {
            PADDING-LEFT: 0;
            PADDING-RIGHT: 0;
            FONT-SIZE: 8pt; 
            TEXT-ALIGN: left;
            BORDER: none;
            }
        .RptGenericCell
            {
            PADDING-LEFT: 0;
            PADDING-RIGHT: 0;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            }
        .RptColHeadCell
            {
            PADDING-LEFT: 20;
            FONT-SIZE: 10pt; 
            TEXT-ALIGN: Left;
            FONT-WEIGHT: bold;
            BORDER-TOP: 1 solid #C0C0C0;
            BORDER-BOTTOM: 1 solid #C0C0C0;
            BACKGROUND:<%=ReqForm("ColBackGround")%>;
            COLOR:<%=ReqForm("ColFontColor")%>;
            }
        .RptHeadingCell
            {
            PADDING-LEFT: 5;
            FONT-SIZE: 10pt; 
            TEXT-ALIGN: left;
            COLOR: <%=ReqForm("SupFontColor")%>;
            BACKGROUND: <%=ReqForm("SupBackGround")%>;
            BORDER-TOP: 1 solid #C0C0C0;
            BORDER-BOTTOM: 1 solid #C0C0C0;
            }
		.ReportText
            {
            PADDING-LEFT: 5;
            PADDING-RIGHT: 5;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: right;
            WIDTH: 75;            
            OVERFLOW: visible;
            }
        .MangementHeading
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            FONT-WEIGHT: bold;
            OVERFLOW: visible;
            WIDTH:75
            }
            
       .ReportTotalsSingle
            {
            PADDING-LEFT: 5;
            PADDING-RIGHT: 5;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            WIDTH: 75
            }
            
        .ReportTotalsDouble
            {
            PADDING-LEFT: 5;
            PADDING-RIGHT: 5;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: double;
            BORDER-TOP-WIDTH: 3;
            BORDER-COLOR:#C0C0C0;
            WIDTH: 75
            }
        .ManagementText
            {
            LEFT:10;
            WIDTH:200;
            PADDING-LEFT: 5;
            PADDING-RIGHT: 5;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            OVERFLOW: hidden;
            }
            
        .ColumnHeading
            {
            OVERFLOW:hidden;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: center;
            FONT-WEIGHT: bold;
            OVERFLOW: hidden;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            BORDER-BOTTOM-STYLE: solid;
            BORDER-BOTTOM-WIDTH: 1;
            WIDTH:75;
            BACKGROUND:<%=ReqForm("ColBackGround")%>;
            COLOR:<%=ReqForm("ColFontColor")%>
            }
            
        .DirectorHeading
            {
            LEFT:10;
            WIDTH:630;
            FONT-SIZE: 10pt; 
            FONT-WEIGHT: bold;
            HEIGHT: 18;
            TEXT-ALIGN: left;
            OVERFLOW: hidden;
            }
        .OfficeHeading
            {
            LEFT:10;
            WIDTH:630;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            OVERFLOW: hidden;
            }
       .ManagerHeading
            {
            LEFT:10;
            WIDTH:625;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            FONT-WEIGHT: normal;
            TEXT-ALIGN: left;
            OVERFLOW: visible;
            }
		.SupervisorHeading
			{
			LEFT:10;
			WIDTH:610;
			FONT-SIZE: 10pt; 
            FONT-WEIGHT: bold;
			HEIGHT: 18;
			TEXT-ALIGN: left;
            OVERFLOW: visible;
			}
		.WorkerHeading
            {
            LEFT:10;
            WIDTH:200;
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: left;
            OVERFLOW: visible;
            }
            
       .DirectorTotals
            {
            LEFT:10;
            WIDTH:240;
            FONT-SIZE: 10pt; 
            FONT-WEIGHT: bold;
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            OVERFLOW: visible;            
            }
        .OfficeTotals
            {
			LEFT:10;
            WIDTH:200;
            FONT-SIZE: 10pt;  
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            OVERFLOW: hidden;
            }
        .ManagerTotals
            {
			LEFT:10;
            WIDTH:240;
            FONT-SIZE: 10pt;  
            FONT-WEIGHT: normal;
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            OVERFLOW: visible;
            }
        .SupervisorTotals
            {
            LEFT:10;
            WIDTH:240;
            FONT-SIZE: 10pt; 
            FONT-WEIGHT: bold;
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            WIDTH:240;
            OVERFLOW: visible;
            }
         .WorkerTotals
            {
            FONT-SIZE: 10pt; 
            HEIGHT: 18;
            TEXT-ALIGN: right;
            BORDER-TOP-STYLE: solid;
            BORDER-TOP-WIDTH: 1;
            BORDER-COLOR:#C0C0C0;
            WIDTH:240;
            OVERFLOW: visible;
            }
         .RptHeader
            {
            BORDER-STYLE:solid; 
            BORDER-WIDTH:1; 
            BORDER-COLOR:#C0C0C0; 
            HEIGHT:75; 
            WIDTH:650; 
            TOP:40; 
            LEFT:10
            }
    </STYLE>