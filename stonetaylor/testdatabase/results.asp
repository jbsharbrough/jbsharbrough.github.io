<%@LANGUAGE="VBSCRIPT"%>

<%
MM_connTravelPackages_STRING = "dsn=compass;"
Dim rsResults__formCat
rsResults__formCat = "%"
if (Request("formCat")   <> "") then rsResults__formCat = Request("formCat")  
%>
<%
Dim rsResults__formPrice
rsResults__formPrice = "100000000"
if (Request("formPrice")     <> "") then rsResults__formPrice = Request("formPrice")    
%>
<%
Dim rsResults__formKey
rsResults__formKey = "%"
if (Request("formKey")  <> "") then rsResults__formKey = Request("formKey") 
%>
<%
set rsResults = Server.CreateObject("ADODB.Recordset")
rsResults.ActiveConnection = MM_connTravelPackages_STRING
rsResults.Source = "SELECT *  FROM TRIPS, PROD_CATEGORIES  WHERE EVENTTYPE LIKE '%" + Replace(rsResults__formCat, "'", "''") + "%' AND PRICE <= " + Replace(rsResults__formPrice, "'", "''") + " AND TRIPDESCRIPTION LIKE '%" + Replace(rsResults__formKey, "'", "''") + "%' AND EVENTTYPE = CATID"
rsResults.CursorType = 0
rsResults.CursorLocation = 2
rsResults.LockType = 3
rsResults.Open()
rsResults_numRows = 0
%>
<%
Dim Repeat1__numRows
Repeat1__numRows = 4
Dim Repeat1__index
Repeat1__index = 0
rsResults_numRows = rsResults_numRows + Repeat1__numRows
%>
<%
'  *** Recordset Stats, Move To Record, and Go To Record: declare stats variables

' set the record count
rsResults_total = rsResults.RecordCount

' set the number of rows displayed on this page
If (rsResults_numRows < 0) Then
  rsResults_numRows = rsResults_total
Elseif (rsResults_numRows = 0) Then
  rsResults_numRows = 1
End If

' set the first and last displayed record
rsResults_first = 1
rsResults_last  = rsResults_first + rsResults_numRows - 1

' if we have the correct record count, check the other stats
If (rsResults_total <> -1) Then
  If (rsResults_first > rsResults_total) Then rsResults_first = rsResults_total
  If (rsResults_last > rsResults_total) Then rsResults_last = rsResults_total
  If (rsResults_numRows > rsResults_total) Then rsResults_numRows = rsResults_total
End If
%>
<%
' *** Recordset Stats: if we don't know the record count, manually count them

If (rsResults_total = -1) Then

  ' count the total records by iterating through the recordset
  rsResults_total=0
  While (Not rsResults.EOF)
    rsResults_total = rsResults_total + 1
    rsResults.MoveNext
  Wend

  ' reset the cursor to the beginning
  If (rsResults.CursorType > 0) Then
    rsResults.MoveFirst
  Else
    rsResults.Requery
  End If

  ' set the number of rows displayed on this page
  If (rsResults_numRows < 0 Or rsResults_numRows > rsResults_total) Then
    rsResults_numRows = rsResults_total
  End If

  ' set the first and last displayed record
  rsResults_first = 1
  rsResults_last = rsResults_first + rsResults_numRows - 1
  If (rsResults_first > rsResults_total) Then rsResults_first = rsResults_total
  If (rsResults_last > rsResults_total) Then rsResults_last = rsResults_total

End If
%>
<%
' *** Move To Record and Go To Record: declare variables

Set MM_rs    = rsResults
MM_rsCount   = rsResults_total
MM_size      = rsResults_numRows
MM_uniqueCol = ""
MM_paramName = ""
MM_offset = 0
MM_atTotal = false
MM_paramIsDefined = false
If (MM_paramName <> "") Then
  MM_paramIsDefined = (Request.QueryString(MM_paramName) <> "")
End If
%>
<%
' *** Move To Record: handle 'index' or 'offset' parameter

if (Not MM_paramIsDefined And MM_rsCount <> 0) then

  ' use index parameter if defined, otherwise use offset parameter
  r = Request.QueryString("index")
  If r = "" Then r = Request.QueryString("offset")
  If r <> "" Then MM_offset = Int(r)

  ' if we have a record count, check if we are past the end of the recordset
  If (MM_rsCount <> -1) Then
    If (MM_offset >= MM_rsCount Or MM_offset = -1) Then  ' past end or move last
      If ((MM_rsCount Mod MM_size) > 0) Then         ' last page not a full repeat region
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' move the cursor to the selected record
  i = 0
  While ((Not MM_rs.EOF) And (i < MM_offset Or MM_offset = -1))
    MM_rs.MoveNext
    i = i + 1
  Wend
  If (MM_rs.EOF) Then MM_offset = i  ' set MM_offset to the last possible record

End If
%>
<%
' *** Move To Record: if we dont know the record count, check the display range

If (MM_rsCount = -1) Then

  ' walk to the end of the display range for this page
  i = MM_offset
  While (Not MM_rs.EOF And (MM_size < 0 Or i < MM_offset + MM_size))
    MM_rs.MoveNext
    i = i + 1
  Wend

  ' if we walked off the end of the recordset, set MM_rsCount and MM_size
  If (MM_rs.EOF) Then
    MM_rsCount = i
    If (MM_size < 0 Or MM_size > MM_rsCount) Then MM_size = MM_rsCount
  End If

  ' if we walked off the end, set the offset based on page size
  If (MM_rs.EOF And Not MM_paramIsDefined) Then
    If (MM_offset > MM_rsCount - MM_size Or MM_offset = -1) Then
      If ((MM_rsCount Mod MM_size) > 0) Then
        MM_offset = MM_rsCount - (MM_rsCount Mod MM_size)
      Else
        MM_offset = MM_rsCount - MM_size
      End If
    End If
  End If

  ' reset the cursor to the beginning
  If (MM_rs.CursorType > 0) Then
    MM_rs.MoveFirst
  Else
    MM_rs.Requery
  End If

  ' move the cursor to the selected record
  i = 0
  While (Not MM_rs.EOF And i < MM_offset)
    MM_rs.MoveNext
    i = i + 1
  Wend
End If
%>
<%
' *** Move To Record: update recordset stats

' set the first and last displayed record
rsResults_first = MM_offset + 1
rsResults_last  = MM_offset + MM_size
If (MM_rsCount <> -1) Then
  If (rsResults_first > MM_rsCount) Then rsResults_first = MM_rsCount
  If (rsResults_last > MM_rsCount) Then rsResults_last = MM_rsCount
End If

' set the boolean used by hide region to check if we are on the last record
MM_atTotal = (MM_rsCount <> -1 And MM_offset + MM_size >= MM_rsCount)
%>
<%
' *** Go To Record and Move To Record: create strings for maintaining URL and Form parameters

' create the list of parameters which should not be maintained
MM_removeList = "&index="
If (MM_paramName <> "") Then MM_removeList = MM_removeList & "&" & MM_paramName & "="
MM_keepURL="":MM_keepForm="":MM_keepBoth="":MM_keepNone=""

' add the URL parameters to the MM_keepURL string
For Each Item In Request.QueryString
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepURL = MM_keepURL & NextItem & Server.URLencode(Request.QueryString(Item))
  End If
Next

' add the Form variables to the MM_keepForm string
For Each Item In Request.Form
  NextItem = "&" & Item & "="
  If (InStr(1,MM_removeList,NextItem,1) = 0) Then
    MM_keepForm = MM_keepForm & NextItem & Server.URLencode(Request.Form(Item))
  End If
Next

' create the Form + URL string and remove the intial '&' from each of the strings
MM_keepBoth = MM_keepURL & MM_keepForm
if (MM_keepBoth <> "") Then MM_keepBoth = Right(MM_keepBoth, Len(MM_keepBoth) - 1)
if (MM_keepURL <> "")  Then MM_keepURL  = Right(MM_keepURL, Len(MM_keepURL) - 1)
if (MM_keepForm <> "") Then MM_keepForm = Right(MM_keepForm, Len(MM_keepForm) - 1)

' a utility function used for adding additional parameters to these strings
Function MM_joinChar(firstItem)
  If (firstItem <> "") Then
    MM_joinChar = "&"
  Else
    MM_joinChar = ""
  End If
End Function
%>
<%
' *** Move To Record: set the strings for the first, last, next, and previous links

MM_keepMove = MM_keepBoth
MM_moveParam = "index"

' if the page has a repeated region, remove 'offset' from the maintained parameters
If (MM_size > 0) Then
  MM_moveParam = "offset"
  If (MM_keepMove <> "") Then
    params = Split(MM_keepMove, "&")
    MM_keepMove = ""
    For i = 0 To UBound(params)
      nextItem = Left(params(i), InStr(params(i),"=") - 1)
      If (StrComp(nextItem,MM_moveParam,1) <> 0) Then
        MM_keepMove = MM_keepMove & "&" & params(i)
      End If
    Next
    If (MM_keepMove <> "") Then
      MM_keepMove = Right(MM_keepMove, Len(MM_keepMove) - 1)
    End If
  End If
End If

' set the strings for the move to links
If (MM_keepMove <> "") Then MM_keepMove = MM_keepMove & "&"
urlStr = Request.ServerVariables("URL") & "?" & MM_keepMove & MM_moveParam & "="
MM_moveFirst = urlStr & "0"
MM_moveLast  = urlStr & "-1"
MM_moveNext  = urlStr & Cstr(MM_offset + MM_size)
prev = MM_offset - MM_size
If (prev < 0) Then prev = 0
MM_movePrev  = urlStr & Cstr(prev)
%>
<html>
<head>
<title>Search Results</title>
<link rel="stylesheet" href="../cart/master.css">
<style type="text/css">
<!--
.background {  background-repeat: no-repeat; background-image:  url(../cart/images/trip_browresults.gif)}
-->
</style>
</head>
<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0 marginwidth=0 marginheight=0 background="../cart/images/trip_browresults.gif" onLoad="" class="background">
<table width="857" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td width="857" height="96"></td>
  </tr>
</table>
<table width="538" border="0" cellpadding="0" cellspacing="0" mm:layoutgroup="true">
  <% 
While ((Repeat1__numRows <> 0) AND (NOT rsResults.EOF)) 
%>
  <tr> 
    <td width="39" height="84" valign="top" class="listname">&nbsp;</td>
    <td width="433" height="84" valign="top" bgcolor="#FFFFCC" class="listname"> 
      <p><a href="../cont_mgmt_cat/detail.asp?id=<%=(rsResults.Fields.Item("TRIPID").Value)%>"><%=(rsResults.Fields.Item("TRIPNAME").Value)%></a></p>
      <p class="normaltext"><%=(rsResults.Fields.Item("TRIPLOCATION").Value)%></p>
      <p class="normaltext"><span class="form">Price: $<%=(rsResults.Fields.Item("PRICE").Value)%><img src="../images/newmemberlogin_imagetop.gif" width="332" height="1"></span><span class="recordcount"> 
        </span></p>
    </td>
    
    <td width="66" height="84" valign="top" bgcolor="#FFFFCC" class="listname">
      <p><img src="images/<%=(rsResults.Fields.Item("catID").Value)%>.gif"></p>
      <p class="recordcount"><%=(rsResults.Fields.Item("catName").Value)%></p>
    </td>
  </tr>
  <tr height=8> 
  </tr>
  <% 
  Repeat1__index=Repeat1__index+1
  Repeat1__numRows=Repeat1__numRows-1
  rsResults.MoveNext()
Wend
%>
  <tr> 
    <td height="553" valign="top" class="listname" colspan="3"> 
      <div align="left" class="recordcount">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=(rsResults_first)%> to <%=(rsResults_last)%> of <%=(rsResults_total)%> 
          


       <object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0" width="93" height="33">
          <param name=movie value="button3.swf">
          <param name=quality value=high>
          <param name="BASE" value=".">
          <param name="BGCOLOR" value="">
          <embed src="button3.swf" base="."  quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="93" height="33" bgcolor="">
          </embed> 
        </object> </div>
      <table border="0" width="50%" align="center">
        <tr> 
          <td width="23%" align="center" height="25"> 
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_moveFirst%>"><img src="../cart/First.gif" border=0></a> 
            <% End If ' end MM_offset <> 0 %>
          </td>
          <td width="31%" align="center" height="25"> 
            <% If MM_offset <> 0 Then %>
            <a href="<%=MM_movePrev%>"><img src="../cart/Previous.gif" border=0></a> 
            <% End If ' end MM_offset <> 0 %>
          </td>
          <td width="23%" align="center" height="25"> 
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveNext%>"><img src="../cart/Next.gif" border=0></a> 
            <% End If ' end Not MM_atTotal %>
          </td>
          <td width="23%" align="center" height="25"> 
            <% If Not MM_atTotal Then %>
            <a href="<%=MM_moveLast%>"><img src="../cart/Last.gif" border=0></a> 
            <% End If ' end Not MM_atTotal %>



          </td>
<td>

</td>
        </tr>
      </table>
    </td>
  </tr>
</table>

</body>
</html>
<%
rsResults.Close()
%>
