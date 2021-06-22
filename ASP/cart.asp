<%Option Explicit%>
<head>
<meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
</head>
<%
    Dim name
    name=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=name
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
	Dim conn
	set conn=server.CreateObject("ADODB.Connection")
	conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
	Dim rs, strsql
	Set rs = Server.CreateObject("ADODB.Recordset")
	Dim rss
	Set rss = Server.CreateObject("ADODB.Recordset")
	rss.Open "Select * From Cart Where Guest like'"&name&"'", conn, 1, 3
%>
<div align="center">
<table width="53%" id="table1">
<th>GName</th><th>Price</th><th>num</th><th>Total</th>
<%Dim i, Totalprice, before,j
  Totalprice=0
  before = 0
  rss.PageSize=1000
  For i=0 To rss.PageSize
    if rss.eof Then Exit For
%>
 <tr>
 		<td width="200"><%=rss("GName")%></td>
		<td><%=rss("Price")%></td>
		<td><%=rss("num")%></td>
		<td><%=rss("Total")%></td>
	</tr>
	<%
	strsql = "Select * From Goods Where GName like '"&rss("GName")&"'"
	rs.Open strsql, conn, 1, 1
	before = before + rss("num") * rs("Price")
	Totalprice = Totalprice + rss("Total")
	rss.MoveNext
	'exit for
	rs.close
	Next
	%>
<title>cart</title>
 <body>
	<h2>
	TotalPrice:￥<%=Totalprice%><br>
	<%if before <> TotalPrice then%>Discount: ￥<%=before-TotalPrice%><%end if%>
	<p align="center"><a href="shopshow.asp">return to shopshow</a></p><br>
	<p align="center"><a href="delete.asp">clear cart</a></p>
	<p align="center"><a href="purchase.asp">purchase</a><p>
 </body>