<%Option Explicit%>
<title>success</title>
<head>
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
    Dim name, strsql
    name=Request.Cookies("user")("name")
    Response.Cookies("user")("name")=name
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())

    strsql="update Users set VIP = 1 where name = '"&name&"'"
	set rs = conn.Execute(strsql)
    Response.Write "购买成功并已成为会员"
%>
<form action="shopshow.asp" method="post">
<input type="submit" name="返回">
</form>
