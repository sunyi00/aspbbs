<%
  set conn=server.createobject("adodb.connection")
  conn.open "DRIVER={Microsoft access driver (*.mdb)};dbq="&server.mappath("bbs.mdb")
%>
