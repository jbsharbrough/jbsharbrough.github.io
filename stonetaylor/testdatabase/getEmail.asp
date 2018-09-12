<%Session("tripid") = Request("id")%>
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#FFFFFF" text="#000000" background="../cart/images/back_getmail.gif">
<div id="Layer1" style="position:absolute; left:62px; top:132px; width:452px; height:122px; z-index:1">
  <form name="form1" method="post" action="sendmail.asp">
    <b><font face="Arial, Helvetica, sans-serif" size="2">enter your email address 
    here: </font> </b> 
    <input type="text" name="email" size="40">
    <input type="submit" name="Submit" value="&gt;&gt;">
    <b><font face="Arial, Helvetica, sans-serif" size="2">
    <input type="hidden" name="id" value="<%=Request("id")%>">
    </font></b> 
  </form>
</div>
</body>
</html>
