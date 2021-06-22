<%Option Explicit%>
 <% Response.Buffer=True %>
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>forget password</title>
 </head>
 <%
    dim name, ans
    name=Request.Cookies("user")("verifyname")
    ans=Request.Cookies("user")("ans")
    
    name=Request("verifyname")
    ans=Request("ans")

    Response.Cookies("user")("ans")=ans
    Response.Cookies("user")("verifyname")=name
    Response.Cookies("user").Expires=DateAdd("d", 30, Date())
 %>
 <%
    Dim conn, rs
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From Users", conn, 1, 1
%>
<%
    dim i
    for i=1 to rs.PageSize
        if rs.Eof then exit for
        if rs("name")=name then
            if rs("answer")=ans then
                Response.write "密码为"+rs("password")%>
                <form action="login.asp" method="post">
                    <input type="submit" name="返回"></form>
            <%else Response.Write "答案不正确"
        end if
    end if
        rs.MoveNext
    next
%>