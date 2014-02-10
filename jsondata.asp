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
	Dim page ,rows ,sidx,sord,sql,ts,t,table
	 page=Easp.R("page",1)
	 rows=Easp.R("rows",1)
	 sidx=Easp.R("sidx",0)
	 sord=Easp.R("sord",0)
	 table=Easp.R("t",0)
	if(sord="desc")then 
		sql="select top "&rows&" * from (select top "&rows*page&"  * from "&table&" order by    "&sidx&" desc) order by "&sidx&" asc" 
	 else 
	sql="select top "&rows&" * from (select top "&rows*page&"  * from "&table&" order by    "&sidx&" asc) order by "&sidx&" desc " 
	end if
	
	
	Dim o
	Set o = jsObject()
	Set o("root") = QueryToJSON(Easp.db.dbConn, sql)
	o("page")=page
	t=Easp.db.ReadTable(table,"1=1", "count(id)")
	dim total
	if (isnumeric(t)) then
        total=((cint((t-1)/rows)))+1
	else
		total=0  
	end if
	o("total")=total
	o("totalSize")=t
	o.Flush  
%>