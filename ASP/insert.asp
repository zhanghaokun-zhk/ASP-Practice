<%@language="VBScript"%>
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
	Dim rs, strsql, rsss
	Set rs = Server.CreateObject("ADODB.Recordset")
	Set rsss = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From Goods", conn, 1, 1
	rsss.Open "Select * From Users", conn, 1, 1
	Dim rss
	Set rss = Server.CreateObject("ADODB.Recordset")
	rss.Open "Select * From Cart", conn, 1, 3
%>
<%
	dim ID, num, Ctotal, Count, No, Total, i
	ID=Trim(Request.QueryString("ID"))
	If ID="" Then 
		Response.End
	End If
	ID=Cint(ID)
	strsql="Select * From Goods Where Gid Like '%"&ID&"%'"
	Set rs=conn.Execute(strsql)
	num=Trim(Request.Form("Text"&ID))
	If num="" Then 
		Response.Redirect("shopshow.asp")
	End If
	num=Cint(num)
	Session("all")=Session("all")+num
	Count=Session("Count")
	No=Session("No")
	Total=Session("Total")
	Session("Count")=Session("Count")+1
	Session("No")=No
	Session("Total")=Total
	rss.addnew
	rss("Guest")=name
	rss("GName")=rs("GName")
	if rsss("name")=name then
	  Ctotal=num*rs("price")*rs("Rebate")
	  rss("price")=rs("price")*rs("Rebate")
	else
	  Ctotal=num*cint(rs("price"))
	  rss("price")=rs("price")
	end if
	rss("num")=num
	rss("Total")=Ctotal
	rss.update
 %>
 <html>
 <head>
	<title>insert</title>
 </head>
 <body>
 <p align="center">success!</p><br><br><br>
 <p align="center"><a href="shopshow.asp">continue</a></p><br><br>
 <p align="center"><a href="cart.asp">check</a></p>
 </body>
</html>