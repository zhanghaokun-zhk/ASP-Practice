<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%Option Explicit%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312" >
</head>
<%
	dim i, ID, strsql
	ID=cint(Trim(Request.QueryString("ID")))
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM Users", conn, 1, 3
%>
<%
	strsql="DELETE FROM Users WHERE ID = "&ID&""
	Response.Write strsql
	set rs = conn.Execute(strsql)
	response.Redirect "admin.asp"
%>