 <%@language="VBSCRIPT"%>
<%Option Explicit%>
<head>
	<title>Goods</title>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Goods", conn, 1, 3
%>
	<center>
		<a href="admin.asp">Back</a>
		<a href="addGoods.asp">Add</a>
<%Dim i
rs.PageSize=1000
  For i=1 To rs.PageSize
    if rs.Eof Then Exit For
%>
<table width="53%" id="table1">
	<th>Gid</th><th>pic</th><th>GName</th><th>Price</th><th>Rebate</th>
<tr>
 		<td><%=rs("Gid")%></td>
 		<td><a href=<%=rs("link")%>><img src="<%=rs("pic")%>" width="200px" height="200px"></td> 
 		<td width="200"><%=rs("GName")%></td>
		<td>￥<%=rs("Price")%></td>
		<td><%=rs("Rebate")%></td>
		<form method="post" action="editgoods.asp?id=<%=rs("Gid")%>">
		<td><input type=submit value="Edit"></td></form>
		<form method="post" action="deletegoods.asp?id=<%=rs("Gid")%>">
		<td><input type=submit value="Delete"></td></form>
	</tr>
</table>
<%
	rs.MoveNext
	next
%>
<nav>
<a href="admin.asp">Back</a>
<a href="addGoods.asp">Add</a>
</nav>
