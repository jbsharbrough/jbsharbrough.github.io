<%@LANGUAGE="VBScript"%>
<HTML>
<!--#include file ="MMHTTPDB.js"-->
<%
On Error Resume Next
Set oConn = CreateMMConnection(Request("ConnectionString"),Request("UserName"),Request("Password"),Request("Timeout"))
Dim opCode
opCode = CStr(Request("opCode"))
if (Len(opCode) AND (opCode <> "undefined")) then
	if (opCode = "GetODBCDSNs") then
		Response.Write(oConn.GetODBCDSNs())
	elseif (opCode = "GetTables") then
		oConn.Open()
		Response.Write(oConn.GetTables(Request("Schema"),Request("Catalog")))
	elseif (opCode = "GetViews") then
		oConn.Open()
		Response.Write(oConn.GetViews(Request("Schema"),Request("Catalog")))
	elseif (opCode = "GetProcedures") then
		oConn.Open()
		Response.Write(oConn.GetProcedures(Request("Schema"),Request("Catalog")))
	elseif (opCode = "GetColsOfTable") then
		oConn.Open()
		Response.Write(oConn.GetColumnsOfTable(Request("TableName"),Request("Schema"),Request("Catalog")))
	elseif (opCode = "GetParametersOfProcedure") then
		oConn.Open()
		Response.Write(MarshallRecordsetIntoHTML(oConn.GetParametersOfProcedure(Request("ProcName"),Request("Schema"),Request("Catalog"))))
	elseif (opCode = "ExecuteSQL") then
		oConn.Open()
		Response.Write(oConn.ExecuteSQL(Request("SQL"),Request("MaxRows")))
	elseif (opCode = "ExecuteSP") then
		oConn.Open()
		Response.Write(oConn.ExecuteSP(Request("ExecProcName"),0,Request("ExecProcParameters")))
	elseif (opCode = "ReturnsResultset") then
		oConn.Open()
		Response.Write(oConn.ReturnsResultSet(Request("RRProcName"),Request("Schema"),Request("Catalog")))
	elseif (opCode = "SupportsProcedure") then
		oConn.Open()
		Response.Write(oConn.SupportsProcedure())
	elseif (opCode =  "GetProviderTypes") then
		oConn.Open()
		Response.Write(oConn.GetProviderTypes())
	elseif (opCode = "IsOpen") then
		oConn.Open()
		Response.Write(oConn.TestOpen())
	end if

	if Err then
		Response.Write(oConn.HandleExceptions())
	end if

	oConn.Close()
end if

%>
</HTML>
