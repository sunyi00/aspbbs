<%@language=vbscript%>
<!--#include file=conn.asp-->

<html>
<head>
</head>
<body bgcolor=#1cf7d8>
<center>
<%
dim strarticletitle,strarticlecontent,strarticleauthor,strarticleid
dim strtable

if session("name")="" then
  response.write "login to publish articles"
  response.end
end if

strarticletitle=request.form("title")
strarticlecontent=request.form("content")
strarticleauthor=session("name")
strarticleid=request.form("articleid")
strtable="article"

if trim(strarticletitle)="" then
  response.write "empty title."
  response.end
end if
if trim(strarticlecontent)="" then
  strarticletitle=strarticletitle & "(empty)"
end if

set rs=server.createobject("adodb.recordset")
rs.open strtable,conn,3,2

rs.addnew
if request.form("submit")="publish" then
  rs("articletitle")=strarticletitle
  rs("articleauthor")=strarticleauthor
  rs("articlecontent")=strarticlecontent
  response.write "article published."
elseif request.form("submit")="fellow" then
  rs("articletitle")=strarticletitle
  rs("articleauthor")=strarticleauthor
  rs("articlecontent")=strarticlecontent
  rs("articleparent")=strarticleid
end if
rs.update
rs.close
set rs=nothing
%>

<%
if request.form("submit")="fellow" then
  strchangesql="update article set articlefellownumber=articlefellownumber+1 where articleid=" & strarticleid
  set connnn=server.createobject("adodb.connection")
  connnn.open conn
  set rs=common.execute(strchangesql)
  set rs=nothing
  conn.close
  set connnn=nothing
  response.write "article fellowed."
end if
%>
</body>
</html>
