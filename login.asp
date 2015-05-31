<%@language=vbscript%>
<!--#include file=conn.asp-->
<%Response.buffer=true%>
<html>
<body bgcolor=#c1f7d8>
<center>

<%
  set rstemp=server.createobject("adodb.recordset")
  dim strname,strpassword,sql
  strname=Request.form("name")
  strpassword=Request.form("password")
  sql="select * from user where username=" & strname & "'"
  rstemp.open sql,conn,1,3
  if rstemp.RecordCount=1 then
    session("name")=strname
    Response.redirect "user.htm"
  else
%>
<a href=javascript:history.back()>login failed, please retry.</a>
<%
  end if
  Response.end
  rstemp.close
  set rstemp=nothing
%>

</center>
</body>
</html>
