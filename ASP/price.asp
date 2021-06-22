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
	dim Price, Rebate
	Price=Trim(Request.Form("Price"&ID))
	Rebate=Trim(Request.Form("Rebate"&ID))
	strsql="update Goods set Price = '"&Price&"' where Gid = '"&cint(ID)&"'"
	set rs = conn.Execute(strsql)
	strsql="update Goods set Rebate = '"&Rebate&"' where Gid = '"&cint(ID)&"'"
	set rs = conn.Execute(strsql)
	Response.redirect "admingoods.asp"
%>