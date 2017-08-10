<%
Sub WriteTableCell(strWidth, strName, strContents)
    Response.Write "<TD ID=" & strName & " class=TableDetail style=""width:" & strWidth & """>" & strContents & "</TD>"
End Sub
%>