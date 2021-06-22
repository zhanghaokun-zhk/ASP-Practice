<%Option Explicit%>
 <% Response.Buffer=True %>
  <%
    Dim conn, rs, strsql, Guest
    set conn=server.CreateObject("ADODB.Connection")
    conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="& Server.MapPath("store.mdb")
    Set rs = Server.CreateObject("ADODB.Recordset")
    Guest=Request("user")("Guest")
    strsql = "delete from Cart where Guest = '"&Guest&"'"
    rs.Open strsql, conn, 1, 1
    Response.redirect "admincart.asp"
%>