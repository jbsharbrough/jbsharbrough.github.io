<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="../Connections/mailingtestconn.asp" -->
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_mailingtestconn_STRING
Recordset1.Source = "SELECT * FROM MAILING "
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0
%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000">
</body>
</html>
<%
Recordset1.Close()
%>
