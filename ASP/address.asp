<%Option Explicit%>
<title>about me</title>
<%
    Dim address, name
    name=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=name
    Response.Cookies("user")("address")=""
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Users Where name like'"&name&"'", conn, 1, 3
%>
<div align="center">
<table width="53%" id="table">
<th>name</th><th>address</th>
<%Dim i
  For i=1 To rs.PageSize
    if rs.Eof Then Exit For
%>
<tr>
 		<td><%=rs("name")%></td>
		<td><%=rs("address")%></td>
		<form method="post" action="insertaddress.asp">
		<td width="70"><input type=text name="address" size="8"></td>
		<td width="70"><input type=submit name="submit" value="Submit"></td>
</tr>
<%
	rs.MoveNext
	next
%>