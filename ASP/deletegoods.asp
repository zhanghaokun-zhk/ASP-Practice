<%Option Explicit%>
<%
	dim i, ID, strsql
	ID=cint(Trim(Request.QueryString("ID")))
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "SELECT * FROM Goods", conn, 1, 3
%>
<%
	strsql="delete from Goods where Gid = '"&cint(ID)&"'"
	set rs = conn.Execute(strsql)
	response.Redirect "admingoods.asp"
%>