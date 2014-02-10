<!--#include file="lib/easp.asp"--> 
<!--#include file="lib/JSON_2.0.4.asp"-->
<!--#include file="lib/JSON_UTIL_0.1.1.asp"-->
<% Session.CodePage=65001
Response.Charset="utf-8" %>
<% 
response.buffer=True
Easp.db.dbConn =Easp.db.OpenConn(1,"db.mdb",""  ) 
%> 



