<%@LANGUAGE="VBSCRIPT"%> 

<%
MM_connTravelPackages_STRING = "dsn=compass;"
email = request("email")
If Request.Cookies("email")="" Then
 Response.Cookies("email")= email
Else 
 email = Request.Cookies("email")
End If 

Dim Recordset1__varID
Recordset1__varID = "1"
if (Request("id") <> "") then Recordset1__varID = Request("id")
%>
<%
set Recordset1 = Server.CreateObject("ADODB.Recordset")
Recordset1.ActiveConnection = MM_connTravelPackages_STRING
Recordset1.Source = "SELECT *  FROM TRIPS  WHERE TRIPID = " + Replace(Recordset1__varID, "'", "''") + ""
Recordset1.CursorType = 0
Recordset1.CursorLocation = 2
Recordset1.LockType = 3
Recordset1.Open()
Recordset1_numRows = 0

Dim the_subject, the_description
the_subject = "Compass Travel: " & Recordset1("TRIPNAME") 
the_partial = Recordset1("TRIPDESCRIPTION")
the_description = the_partial & vbCRLF & vbCRLF &_
               "Price: $" & Recordset1("Price") & vbCRLF & vbCRLF &_
               "For more information, please contact us at info@compasstravel.com."

%>
<% 
Dim objCDO
Set objCDO = Server.CreateObject("CDONTS.NewMail")
objCDO.From = "info@compasstravel.com"
objCDO.To = email
objCDO.CC = ""
objCDO.Subject = the_subject
objCDO.Body = the_description
objCDO.Send()
Set objCDO = Nothing







%>
<html>
<head>
<title>Your email has been sent! </title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>
<body bgcolor="#FFFFFF" text="#000000" background="../cart/images/back_sent.gif">
 
<div id="Layer1" style="position:absolute; left:55px; top:175px; width:345px; height:305px; z-index:1">
  <p><font face="Arial, Helvetica, sans-serif" size="3"><b><font color="#000000">Thank 
    you, <%=email%>!</font></b></font></p>
  <p><font face="Arial, Helvetica, sans-serif" size="3" color="#993300"><b><%=(Recordset1.Fields.Item("TRIPNAME").Value)%></b></font></p>
  <p><font color="#CC3333" size="2"><b><%=(Recordset1.Fields.Item("TRIPLOCATION").Value)%></b></font></p>
  <p><font face="Arial, Helvetica, sans-serif" size="2" color="#FFFFFF"><b><%=the_partial%></b></font></p>
  <p>&nbsp;</p>
  <p>&nbsp;</p>
  <p><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="93" height="33">
      <param name=movie value="button7.swf">
      <param name=quality value=high>
      <param name="BASE" value=".">
      <param name="BGCOLOR" value="#FFCC66">
      <embed src="button7.swf" base="."  quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="93" height="33" bgcolor="#FFCC66">
      </embed> 
    </object></p>
</div>
</body>
</html>
<%
Recordset1.Close()
%>
