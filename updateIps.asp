<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="lib/easp.asp"--> 
<!--#include file="lib/JSON_2.0.4.asp"-->
<!--#include file="lib/JSON_UTIL_0.1.1.asp"-->
<% Session.CodePage=65001
Response.Charset="utf-8" %>
<% 
response.buffer=True
Easp.db.dbConn =Easp.db.OpenConn(1,"db.mdb",""  ) 
%> 
<%  
	dim ip
	ip=getIP() 
%> 
<%

if (Easp.ra("oper",0)="del") then 
	Dim a(1,1) 
	Easp.db.DeleteRecord "ips", "id in(" & Easp.R("id",1) &")"
	 a(0,0) = "result"
	a(0,1) = "success"
	a(1,0) = "message"
	a(1,1) = "删除成功"
	Response.Write toJSON(a)
end if

%>
<%

if (Easp.ra("oper",0)="edit") then 
 	Dim result   
	result = Easp.db.UpdateRecord("ips","id="&Easp.R("id",1),Array("name:"&Easp.ra("name",0),"ip:"&ip)) 
end if

%>
<%

if (Easp.ra("oper",0)="add") then 
		Dim resultadd
		resultadd=  Easp.db.AddRecord("ips",Array("name:"&Easp.ra("name",0),"ip:"&ip)) 
	 
end if

%>

<%
Private Function getIP()   
Dim strIPAddr   
If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then   
strIPAddr = Request.ServerVariables("REMOTE_ADDR")   
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then   
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)   
ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then   
strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)   
Else   
strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")   
End If   
getIP = Trim(Mid(strIPAddr, 1, 30))   
End Function    
%>