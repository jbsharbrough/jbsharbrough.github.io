<%@LANGUAGE="VBSCRIPT"%> 
<%
id = Request("TRIPID")
If Request("sendPressed") = "true" Then
 If Request.Cookies("email") <> "" Then
  redir_page = "sendmail.asp?id=" & id
  Response.Redirect(redir_page)
 Else
  redir_page = "getEmail.asp?id=" & id
  Response.Redirect(redir_page)
 End If 
End If

MM_connTravelPackages_STRING = "dsn=compass;"
Dim rsDetail__varID
rsDetail__varID = "1"
if (Request("id")  <> "") then rsDetail__varID = Request("id") 
%>
<%
set rsDetail = Server.CreateObject("ADODB.Recordset")
rsDetail.ActiveConnection = MM_connTravelPackages_STRING
rsDetail.Source = "SELECT *  FROM TRIPS  WHERE TRIPID = " + Replace(rsDetail__varID, "'", "''") + ""
rsDetail.CursorType = 0
rsDetail.CursorLocation = 2
rsDetail.LockType = 3
rsDetail.Open()
rsDetail_numRows = 0
%>
<%
set rsCart = Server.CreateObject("ADODB.Recordset")
rsCart.ActiveConnection = MM_connTravelPackages_STRING
rsCart.Source = "SELECT *  FROM CART"
rsCart.CursorType = 0
rsCart.CursorLocation = 2
rsCart.LockType = 3
rsCart.Open()
rsCart_numRows = 0
%>
<%
dim redir_page
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsResults_numRows = rsResults_numRows + Repeat1__numRows

' mbarbarelli -- Code Mod
%>

<html>
<head>
<title>Details about your journey</title>
<link rel="stylesheet" href="../cart/master.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 background="../cart/images/trip_details.gif" onLoad="">
<table width="564" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="64" height="128"></td>
    <td width="440"></td>
    <td width="60"></td>
  </tr>
  <tr> 
    <td height="227"></td>
    <td valign="top" height="227"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="69" height="50" valign="top"><img src="../images/whiteWaterRafting_f5.gif" width="65" height="49"> 
          </td>
          <td width="371" height="50" class="listname"><%=(rsDetail.Fields.Item("TRIPNAME").Value)%></td>
        </tr>
        <tr> 
          <td height="130" width="69" rowspan="2"></td>
          <td align="left" valign="top" width="371" class="normaltext" height="104"> 
            <p>Price: <span class="detaillabels"> $<%=(rsDetail.Fields.Item("PRICE").Value)%></span></p>
            <p class="normaltext">&nbsp;</p>
            <p class="normaltext">&nbsp;</p>
            <p class="normaltext">&nbsp;</p>
            <p class="normaltext">&nbsp;</p>
            <p class="normaltext">&nbsp; 
          </td>
        </tr>
        <tr> 
          <td align="left" valign="top" width="371" class="normaltext" height="65"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="93" height="33">
              <param name=movie value="button1.swf">
              <param name=quality value=high>
              <param name="BASE" value=".">
              <param name="BGCOLOR" value="">
              <embed src="button1.swf" base="."  quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="93" height="33" bgcolor="">
              </embed> 
            </object>
<form name="sendinfo">
              <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="93" height="33">
                <param name=movie value="button6.swf">
                <param name=quality value=high>
                <param name="BGCOLOR" value="">
                <embed src="button6.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="93" height="33" bgcolor="" base="">
                </embed> 
              </object>
              <input type="hidden" name="sendPressed" value="true">
              <span class="recordcount"><< send me more info! 
              <input type="hidden" name="TRIPID" value="<%=(rsDetail.Fields.Item("TRIPID").Value)%>">
              </span> 
            </form>
</td>
        </tr>
      </table>
    </td>
    <td height="227"></td>
  </tr>
  <tr> 
    <td height="104"></td>



    <td align="left" valign="top" class="detaillabels">
</td>
    <td></td>
  </tr>
</table>
<p>&nbsp;</p></body>
</html>
<%
rsDetail.Close()
%>
<%
rsCart.Close()
%>
