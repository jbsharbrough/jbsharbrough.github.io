
<!--#include file="../Connections/mailingconn.asp" -->
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_mailingconn_STRING
Recordset1.Source = "SELECT * FROM MAILING"
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<%
Recordset1.Close()
%>

