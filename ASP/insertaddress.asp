<%Option Explicit%>
<title>about me</title>
<%
    Dim name, address
    name=Request.Cookies("user")("name")
    address=Request.Cookies("user")("address")
    Response.Cookies("user")("address")=address
    address=Request("address")
    Response.Cookies("user")("name")=name
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
    Dim conn
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Dim rss, strsql
    Set rss = Server.CreateObject("ADODB.Recordset")
    'rss.Open "Select * From Users Where name like'"&name&"'", conn, 1, 3
    strsql = "update Users set address = '"&address&"' where name = '"&name&"'"
    rss.Open strsql, conn, 1, 3
    Response.redirect "shopshow.asp"
%>