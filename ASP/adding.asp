  <head>
	<title>AddGoods</title>
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
<%
  dim GName, Price, GDScript, Rebate, pic, link
    GName=Request.Cookies("user")("GName")
    Price=Request.Cookies("user")("Price")
    GDScript=Request.Cookies("user")("GDScript")
    Rebate=Request.Cookies("user")("Rebate")
    pic=Request.Cookies("user")("pic")
    link=Request.Cookies("user")("link")

    GName=Request("GName")
    Price=Request("Price")
    GDScript=Request("GDScript")
    Rebate=Request("Rebate")
    pic=Request("pic")
    link=Request("link")

    Response.Cookies("user")("GName")=GName
    Response.Cookies("user")("Price")=Price
    Response.Cookies("user")("GDScript")=GDScript
    Response.Cookies("user")("Rebate")=Rebate
    Response.Cookies("user")("pic")=pic
    Response.Cookies("user")("link")=link
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
%>
<%
  dim flag,Gid
  flag = 0
  rs.PageSize = 1000
  for i = 1 to rs.PageSize
    if rs.Eof then exit for
    if GName = rs("GName") then
      flag = 1
      exit for
    end if
    Gid=cint(rs("Gid"))
    rs.MoveNext
  next
  if flag = 0 then
    rs.addnew
    rs("Gid")=11
    rs("GName")=GName
    rs("Price")=Price
    rs("GDScript")=GDScript
    rs("Rebate")=Rebate
    rs("pic")=pic
    rs("link")=link
    rs.update
    rs.close
  end if
  Response.redirect "admingoods.asp"
%>