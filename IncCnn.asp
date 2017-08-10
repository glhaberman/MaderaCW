<%
'==========================================================================='
' Case Review System - Rushmore Group, LLC                                  '
'---------------------------------------------------------------------------'
'    Name: IncCnn.asp                                                       '
' Purpose: This include is used on all pages in the application.  It is     '
'          responsible for making the database connection, setting any      '
'          IIS config options, and setting variables that are used on all   '
'          pages.                                                           '
'                                                                           '
'==========================================================================='
'-- ADO Type Library:
%>
<!-- METADATA TYPE="TypeLib" UUID="00000205-0000-0010-8000-00AA006D2EA4" -->
<%
'ADO Type Library GUID's - modify the previous METADATA tag with the
'correct type library ID that is installed on the internet server hosting
'the application.  Some common id's are shown below:
'
'  ADO 2.5 Type Library:
'    UUID="00000205-0000-0010-8000-00AA006D2EA4"
'
'  ADO 2.6 Type Library:
'    UUID="00000206-0000-0010-8000-00AA006D2EA4"
'
'  ADO 2.7 Type Library:
'    UUID="EF53050B-882E-4776-B643-EDA472E8E3F2"
'
'  ADO 2.8 Type Library:
'    UUID="2A75196C-D9EB-4129-B803-931327F72D5C"
'
'---------------------------------------------------------------------------'
Dim strDb          'SQL Database Name.
Dim strServer      'SQL Server Name.
Dim strSqlUser     'SQL User ID.
Dim strSqlPW       'SQL Password.

'IIS Settings:
Server.ScriptTimeout = 240
Response.Buffer = True
Response.ExpiresAbsolute = Now - 365
gblnUseLogon = True

'Database Connection Settings:
strServer = "2G2M362GH\SQL2014"
strDb = "MaderaCW"

'If the SQL user ID and password variables are left blank, the connection will
'be attempted assuming SQL Windows authentication is enabled.  If ID and
'password are supplied, SQL authentication will be attempted.:
If gblnUseLogon Then
    strSqlUser = "sa"
    strSqlPW = "TRG123!"
Else
    strSqlUser = ""
    strSqlPW = ""
End If

'Open the database connection:
Call OpenConnection(strServer, strDb, strSqlUser, strSqlPW)
'Initialize global variables and settings used on every page:
Call GetGlobalSettings
%>
<!--#include file="IncGlobals.asp"-->
<!--#include file="IncSvrFunctions.asp"-->
<!--#include file="IncDebug.asp"-->
