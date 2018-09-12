<%@LANGUAGE="VBSCRIPT"%>

<%
MM_connTravelPackages_STRING="DRIVER={Microsoft Access Driver (*.mdb)};DBQ=http://www.stonetaylor.com/testdatabase/compasstravel.mdb 
'MM_connTravelPackages_STRING = "dsn=compass;"
set rsResults = Server.CreateObject("ADODB.Recordset")
rsResults.ActiveConnection = MM_connTravelPackages_STRING
rsResults.Source = "SELECT DISTINCT catName, catID  FROM PROD_CATEGORIES"
rsResults.CursorType = 0
rsResults.CursorLocation = 2
rsResults.LockType = 3
rsResults.Open()
rsResults_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = -1
Dim Repeat1__index
Repeat1__index = 0
rsResults_numRows = rsResults_numRows + Repeat1__numRows
%>
<html>
<head>
<title>Search for your trip</title>
<link rel="stylesheet" href="../cart/master.css">
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 background="../cart/images/trip_cat.gif" onLoad="">
<form action="../cont_mgmt_cat/results.asp" name="form1">
  <table width="519" border="0" cellpadding="0" cellspacing="0">
    <tr> 
      <td width="31" height="105"></td>
      <td width="418" height="105"></td>
      <td width="113" height="105"></td>
  </tr>
  <tr> 

      <td height="195" width="31"></td>
      <td valign="top" width="418"> 
        <table border="0" cellpadding="0" cellspacing="0" width="423">
          <!-- fwtable fwsrc="ultimatesearch.png" fwbase="ultimatesearch.gif" fwstyle="Dreamweaver" fwdocid = "742308039" fwnested="0" -->
          <tr> 
            <td width="186"><img src="../cart/spacer.gif" width="186" height="1" border="0"></td>
            <td width="85"><img src="../cart/spacer.gif" width="51" height="1" border="0"></td>
            <td colspan="2"><img src="../cart/spacer.gif" width="113" height="1" border="0"></td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="1" border="0"></td>
          </tr>
          <tr> 
            <td colspan="4"><img name="ultimatesearch_r1_c1" src="../cart/ultimatesearch_r1_c1.gif" width="350" height="83" border="0"></td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="83" border="0"></td>
          </tr>
          <tr> 
            <td rowspan="7" width="186"><img name="ultimatesearch_r2_c1" src="../cart/ultimatesearch_r2_c1.gif" width="186" height="217" border="0"></td>
            <td colspan="3"> 
              <input type="text" name="formKey" class="search">
            </td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="24" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"><img name="ultimatesearch_r3_c2" src="../cart/ultimatesearch_r3_c2.gif" width="164" height="16" border="0"></td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="16" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <select name="formCat" class="search">
                <option value="%">All</option>
                <%
While (NOT rsResults.EOF)
%>
                <option value="<%=(rsResults.Fields.Item("catID").Value)%>" ><%=(rsResults.Fields.Item("catName").Value)%></option>
                <%
  rsResults.MoveNext()
Wend
If (rsResults.CursorType > 0) Then
  rsResults.MoveFirst
Else
  rsResults.Requery
End If
%>
              </select>
            </td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="24" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"><img name="ultimatesearch_r5_c2" src="../cart/ultimatesearch_r5_c2.gif" width="164" height="17" border="0"></td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="17" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"> 
              <select name="formPrice" class="search">
                <option value="1000000">No Limit</option>
                <option value="200">$200</option>
                <option value="500">$500</option>
                <option value="700">$700</option>
                <option value="900">$900</option>
                <option value="1000">$1000</option>
                <option value="1500">$1500</option>
                <option value="2000">$2000</option>
                <option value="2500">$2500</option>
              </select>
            </td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="25" border="0"></td>
          </tr>
          <tr> 
            <td colspan="3"><img name="ultimatesearch_r7_c2" src="../cart/ultimatesearch_r7_c2.gif" width="164" height="80" border="0"></td>
            <td width="3"><img src="../cart/spacer.gif" width="1" height="80" border="0"></td>
          </tr>
          <tr> 
            <td width="85"><img name="ultimatesearch_r8_c2" src="../cart/ultimatesearch_r8_c2.gif" width="51" height="31" border="0"></td>
            <td width="147"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="123" height="32">
                <param name=movie value="../cart/button5.swf">
                <param name=quality value=high>
                <embed src="../cart/button5.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="123" height="32">
                </embed> 
              </object></td>

            <td width="2"><img src="../cart/spacer.gif" width="1" height="31" border="0"></td>
          </tr>
        </table>
      </td>
      <td width="113" valign="bottom"> 
   </td>
  </tr>
  <tr> 
      <td height="140" width="31"></td>
      <td width="418"></td>
      <td width="113"></td>
  </tr>
</table>
</form>
</body>
</html>
<%
rsResults.Close()
%>
