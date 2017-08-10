<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncTableStyles.asp                                               '
' Purpose: This include file contains styles used to simulate a listview    '
'          with an HTML table.                                              '
'                                                                           '
'          TableDivArea is for the auto scroll <DIV> containing the table.  '
'          CellLabel defines the attributes of header cells in the table.   '
'          TableRow is the style for an unselected row.                     '
'          TableRowSelected is the style for a row selected by the user.    '
'==========================================================================='
%>

    <STYLE id=TableStyles type="text/css" rel="stylesheet">
        .TableDivArea
            {
            BORDER-STYLE: solid;
            BORDER-WIDTH: 1px;
            BORDER-COLOR: <%=gstrBorderColor%>;
            BACKGROUND-color: white;
            POSITION: absolute;
            OVERFLOW: auto
            }

        .CellLabel
            {
            height: 15;
            padding-left: 5px;
            padding-right: 5px;
            border-style: solid;
            border-width: 1px;
            font-family: Tahoma;
            font-size:8pt;
            font-weight: bold;
            text-align: center;
            color:black;
            background-color: <%=gstrPageColor%>;
            overflow: hidden
            }

        .TableRow
            {
            background-color: white;
            color: black;
            }
    
        .TableSelectedRow
            {
            background-color: <%=gstrTitleColor%>;
            color: white;
            }

        .TableDetail
            {
            height: 15;
            font-family: Tahoma;
            font-size:8pt;
            border: 1 solid grey;
            text-align: left;
            padding-left: 5;
            padding-right: 5;
            overflow: hidden
            }
    </STYLE>