<%@language=vbscript%>
<!--#include file=conn.asp-->
<html>
<body bgcolor=#c1f7d8>
<center>

<%
  dim stroldpassword,strnewpassword,strconfirmpassword
  dim strwhere,strsql,strchangesql
  
  stroldpassword=Request.form("oldpassword")
  strnewpassword=Request.form("newpassword")
  strconfirmpassword=Request.form("confirmpassword")

  if stroldpassword="" or strnewpassword="" then
    Response.Write "please input password"
    Response.end
  end if
  if strnewpassword <> strconfirmpassword then
    Response.Write "two new password is not the same"
    Response.end
  end if

  strwhere="where username='" & session("name") & "' and userpassword='" & stroldpassword & "'"
  strsql="select * from user " & strwhere
  strchangesql="update user set userpassword='" & strnewpassword & "' " & strwhere

  set rs=server.createobject("adodb.recordset")
  rs.open strsql,conn,1,3

  if rs.RecordCount=1 then
    set changers=server.createobject("adodb.recordset")
    changers.open strchangesql,conn,1,3
    set changers=nothing
    Response.Write "password changed."
  else
    Response.Write "invalid: password changing failed."
  end if
  rs.close
  set rs=nothing
%>

</center>
</body>
</html>
