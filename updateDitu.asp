<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/easp.asp"--> 
<% Session.CodePage=65001
Response.Charset="utf-8" %>
<% 
response.buffer=True
Easp.db.dbConn =Easp.db.OpenConn(1,"db.mdb",""  ) 
%> 
<%  
	 
%> 
<%

if (Easp.ra("oper",0)="del") then 
	Dim a(1,1) 
	Easp.db.DeleteRecord "map", "id in(" & Easp.R("id",1) &")"
end if

%>
<%

if (Easp.ra("oper",0)="edit") then 
 	Dim result   
	result = Easp.db.UpdateRecord("map","id="&Easp.R("id",1),Array("x:"&Easp.ra("x",0),"y:"&Easp.ra("y",0),"z:"&Easp.ra("z",0),"huangjin:"&Easp.ra("huangjin",0),"jinshu:"&Easp.ra("jinshu",0),"qingqi:"&Easp.ra("qingqi",0),"mucai:"&Easp.ra("mucai",0),"baoshi:"&Easp.ra("baoshi",0),"fuyouziyuan:"&Easp.ra("fuyouziyuan",0),"teseziyuan:"&Easp.ra("teseziyuan",0),"suoyouzhe:"&Easp.ra("suoyouzhe",1))) 
end if

%>
<%

if (Easp.ra("oper",0)="add") then 
	Dim resultadd
	resultadd=  Easp.db.AddRecord("map",Array("x:"&Easp.ra("x",0),"y:"&Easp.ra("y",0),"z:"&Easp.ra("z",0),"huangjin:"&Easp.ra("huangjin",1),"jinshu:"&Easp.ra("jinshu",1),"qingqi:"&Easp.ra("qingqi",0),"mucai:"&Easp.ra("mucai",0),"baoshi:"&Easp.ra("baoshi",0),"fuyouziyuan:"&Easp.ra("fuyouziyuan",0),"teseziyuan:"&Easp.ra("teseziyuan",0),"suoyouzhe:"&Easp.ra("suoyouzhe",1))) 
end if

%>
<%


%>





