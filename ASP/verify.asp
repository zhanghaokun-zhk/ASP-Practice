<%
If Trim(Request.Form("validatecode"))=Empty Or Trim(Session("cnbruce.com_ValidateCode"))<>Trim(Request.Form("validatecode")) Then
            response.write("verify code")
            response.end
            end if
%>