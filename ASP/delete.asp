<%Option Explicit%>
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
    Dim rss
    Set rss = Server.CreateObject("ADODB.Recordset")
    conn.Execute "Delete * From Cart Where Guest like '"&name&"'"
    Response.redirect "cart.asp"
%>
<!DOCTYPE html>
<html lang="en">
 <head>
  <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>Document</title>
 </head>
 <body>
 </body>
</html>
