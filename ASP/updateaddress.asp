<%Option Explicit%>
<?xml version="1.0" encoding="UTF-8"?>
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
  Dim rs, strsql
  Set rs = Server.CreateObject("ADODB.Recordset")
  dim address, name
  address = Request.Cookies("user")("address")
  Response.Cookies("user")("address")=address
  address=Request("address")
  name=Request.Cookies("user")("name")
  Response.Cookies("user")("name")=name
%>
<head>
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />
    <title>修改地址</title>
<body>
      <h2 class="form-signin-heading">修改地址</h2>
      <form action="" method="post">
        地址:<input type="text" name="address">
        <input type="submit" name="确定"></form>
<%
    if address <> "" then 
      strsql = "update Users set address = '"&address&"' where name='"&name&"'"
      rs.Open strsql, conn, 1, 3
      Response.redirect "purchase.asp"
    end if
%>
</body>
</html>