<%Option Explicit%>
 <% Response.Buffer=True %>
 <head>
 <meta charset="UTF-8">
  <meta name="Generator" content="EditPlus®">
  <meta name="Author" content="">
  <meta name="Keywords" content="">
  <meta name="Description" content="">
  <title>admin cart</title>
  <%
    Dim conn, rs, strsql
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Set rs = Server.CreateObject("ADODB.Recordset")
    Response.Cookies("user")("Guest")=""
%>
<%
    dim Guest, i
    Guest=Request("Guest")
    if Request("clear")="delete" then 
        strsql="delete * from Cart Where Guest = '"&Guest&"'"
        'conn.Execute strsql
        'Response.redirect "admincart.asp"
    else
    if Guest="" then 
        strsql="Select * from Cart"
    else
        strsql="Select * from Cart Where Guest = '"&Guest&"'"
    end if
end if
%>
<form action="" method="post">
    <input type="text" name="Guest">
      <input type="submit" value="delete" name="clear">
      <input type="submit" value="确定">
</form>
<table width="53%" id="table1">
<%    
    rs.Open strsql, conn, 1, 3
    rs.PageSize=1000%>
    <th>Guest</th><th>GName</th><th>num</th><th>Total</th>
<%
    for i=1 to rs.PageSize
        if rs.Eof then exit for
%>
<tr>
        <td><%=rs("Guest")%></td>
        <td><%=rs("GName")%></td>
        <td><%=rs("num")%></td>
        <td><%=rs("Total")%></td>
    </tr>
<%
    rs.MoveNext
    next
%>
<nav><a href="admin.asp">back</a></nav>