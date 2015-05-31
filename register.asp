<%@language=vbscript%>
<!--#include file=conn.asp-->
<html>
<head></head>
<body bgcolor=#c1f7d8>
<center>
<%
  dim strname,strpassword1,strpassword2,stremail,strhomepage,strnote,strpassword

  strname=Request.form("name")
  strpassword1=Request.form("password1")
  strpassword2=Request.form("password2")
  stremail=Request.form("email")
  strhomepage=Request.form("homepage")
  strnote=Request.form("note")

  if strname="" then
    Response.Write "invalid: empty UserName<p></p>"
%>
    <a href=javascript:history.back()>back</a>
<%
    Response.End
  end if
  if strpassword1="" then
    Response.Write "invalid: empty UserPassword<p></p>"
%>
    <a href=javascript:history.back()>back</a>
<%
    Response.End
  end if
  if strpassword2="" then
    Response.Write "invalid: empty ConfirmPassword<p></p>"
%>
    <a href=javascript:history.back()>back</a>
<%
    Response.End
  end if
  if strpassword1=strpassword2 then
    strpassword=strpassword1
  else
    Response.Write "invalid: please check two password input<p></p>"
%>
    <a href=javascript:history.back()>back</a>
<%
    Response.End
  end if

  strsql="select * from user where username='" & strname & "'"
  set rs=server.createobject("adodb.recordset")
  rs.open strsql,conn,1,3

  if not (rs.eof and rs.eof) then
    Response.Write "invalid: username already exists, please change username<br><br>"
%>
    <a href=javascript:history.back()>back</a>
<%
    Response.End
  end if
  rs.close
  set rs=nothing


  strtable="user"
  set rs=server.createobject("adodb.recordset")
  rs.open strtable,conn,1,3
  rs.addnew
  rs("username")=strname
  rs("UserPassword")=strpassword
  rs("useremail")=stremail
  rs("userhomepage")=strhomepage
  rs("usernote")=strnote
  rs.update
  rs.close
  set rs=nothing

  Response.Write "Congratulations! Successful Registeration."

%>

<br>

<a href="default.asp">back to index</a>

</body>
</html>
