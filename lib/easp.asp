<%
'#################################################################################
'##	easp.asp
'##	--------------------------------------------------------------------
'##	Feature		:	EasyAsp Class
'##	Version		:	v2.1 beta
'##	Author		:	Coldstone(coldstone[��]qq.com)
'##	Update Date	:	2009-03-26
'##	Description	:	EasyAsp��
'##
'#################################################################################
'
'EasyAsp V2.1 beta ����˵�� [2009-03-26, By Coldstone]
'
'���������ܡ�
'����Easp.regReplaceM���������������滻�Ķ���ģʽ��
'����Easp.regMatch��������������ƥ��ı��鲶��
'����Easp.isInstall���������ڼ��ϵͳ�Ƿ�װ��ĳ�������
'����Easp.Include����������ʵ����ASP�Ķ�̬��������֧��ASPԴ�������޼���ε�<!--#include...-->��
'����Easp.getInclude���������ڻ�ȡASP�ļ����еĽ�����ȡhtml�ļ����ı��ļ���Դ�롣
'����Easp.db.QueryType���ԣ���������RecordSet����Command��ʽ��ȡ��¼����
'����Easp.db.GetRandRecord����������ȡ����������ļ�¼����
'����Easp.db.Exec����������ִ��SQL�����߷���Command��ʽ��ѯ�ļ�¼����
'
'���������¡�
'�Ż�Easp.DateTime��������ʽ��Ϊʱ���ʱ����ʾ�����Ի���
'�Ż�Easp.db.AddRecord���������ڽ���ָ����ID�ֶε�����ʱ�ŷ��������ļ�¼ID�š�
'����EasyAsp V2.0�з�ҳ��������ҳ������С��jumplong����ʱ���ָ�����Bug��

Dim Easp : Set Easp = New EasyASP

Dim EasyAsp_s_html
Class EasyAsp

	Public db
	Private s_path, s_fsoName
	
	Private Sub Class_Initialize()
		s_path		= "/easp/"						'Easp���ļ���·��
		s_fsoName	= "Scripting.FilesyStemObject"	'Ĭ��FSO�������
		Set db		= New EasyAsp_db
	End Sub
	
	Private Sub Class_Terminate()
		Set db 		= Nothing
	End Sub
	
	Public Property Let basePath(ByVal path)
		s_path = path
	End Property
	Public Property Let fsoName(ByVal str)
		s_fsoName = str
	End Property
'***** Easp���������� *****
'����ַ���(���׶ϵ����)
Sub W(ByVal str)
	Response.Write(str)
End Sub
Sub WC(ByVal str)
	Response.Write(str & VbCrLf)
End Sub
Sub WN(ByVal str)
	Response.Write(str & "<br />" & VbCrLf)
End Sub
Sub WE(ByVal str)
	Response.Write(str)
	Response.End()
End Sub
'����������ת
Sub RR(ByVal str)
	Response.Redirect(str)
End Sub
'�ж��Ƿ�Ϊ��ֵ
Function isN(ByVal str)
	isN = Easp_isN(str)
End Function
'�ж���Ԫ���ʽ
Function IIF(ByVal Cn, ByVal T, ByVal F)
	IIF = Easp_IIF(Cn,T,F)
End Function
Function IfThen(ByVal Cn, ByVal T, ByVal F)
	IfThen = Easp_IIF(Cn,T,F)
End Function
'�����������javascript
Sub Js(ByVal Str)
	Response.Write("<sc" & "ript type=""text/javascript"">" & VbCrLf)
	Response.Write(VbTab & Str & VbCrLf)
	Response.Write("</scr" & "ipt>" & VbCrLf)
End Sub
'�����������javascript������Ϣ�򲢷���ǰҳ
Sub Alert(ByVal str)
	Response.Write("<sc" & "ript type=""text/javascript"">alert('" & JsEncode(str) & "\t\t');history.go(-1);</sc" & "ript>"&VbCrLf)
	Response.End()
End Sub
'�����������javascript������Ϣ��ת��URL
Sub AlertUrl(ByVal str, ByVal url)
	Response.Write("<sc" & "ript type=""text/javascript"">"&VbCrLf)
	Response.Write(VbTab&"alert('" & JsEncode(str) & "\t\t');location.href='" & url & "';"&VbCrLf)
	Response.Write("</sc" & "ript>"&VbCrLf)
End Sub
'�����ַ����е�Javascript�����ַ�
Function JsEncode(ByVal str)
	JsEncode = Easp_JsEncode(str)
End Function
'�����ַ�����
Function Escape(ByVal str)
	Escape = Easp_Escape(str)
End Function
'�����ַ�����
Function UnEscape(ByVal str)
	UnEscape = Easp_UnEscape(str)
End Function
'��ʽ������ʱ��
Function DateTime(ByVal iTime, ByVal iFormat)
	If Not IsDate(iTime) Then DateTime = "Date Error" : Exit Function
	If Instr(",0,1,2,3,4,",","&iFormat&",")>0 Then DateTime = FormatDateTime(iTime,iFormat) : Exit Function
	Dim diffs,diffd,diffw,diffm,diffy,dire,before,pastTime
	Dim iYear, iMonth, iDay, iHour, iMinute, iSecond,iWeek,tWeek
	Dim iiYear, iiMonth, iiDay, iiHour, iiMinute, iiSecond,iiWeek
	Dim iiiWeek, iiiMonth, iiiiMonth
	Dim SpecialText, SpecialTextRe,i,t
	iYear = right(Year(iTime),2) : iMonth = Month(iTime) : iDay = Day(iTime)
	iHour = Hour(iTime) : iMinute = Minute(iTime) : iSecond = Second(iTime)
	iiYear = Year(iTime) : iiMonth = right("0"&Month(iTime),2)
	iiDay = right("0"&Day(iTime),2) : iiHour = right("0"&Hour(iTime),2)
	iiMinute = right("0"&Minute(iTime),2) : iiSecond = right("0"&Second(iTime),2)
	tWeek = Weekday(iTime)-1 : iWeek = Array("��","һ","��","��","��","��","��")
	If isDate(iFormat) or isN(iFormat) Then
		If isN(iFormat) Then : iFormat = Now() : pastTime = true : End If
		dire = "��" : If DateDiff("s",iFormat,iTime)<0 Then : dire = "ǰ" : before = True : End If
		diffs = Abs(DateDiff("s",iFormat,iTime))
		diffd = Abs(DateDiff("d",iFormat,iTime))
		diffw = Abs(DateDiff("ww",iFormat,iTime))
		diffm = Abs(DateDiff("m",iFormat,iTime))
		diffy = Abs(DateDiff("yyyy",iFormat,iTime))
		If diffs < 60 Then DateTime = "�ո�" : Exit Function
		If diffs < 1800 Then DateTime = Int(diffs\60) & "����" & dire : Exit Function
		If diffs < 2400 Then DateTime = "��Сʱ"  & dire : Exit Function
		If diffs < 3600 Then DateTime = Int(diffs\60) & "����" & dire : Exit Function
		If diffs < 259200 Then
			If diffd = 3 Then DateTime = "3��" & dire & " " & iiHour & ":" & iiMinute : Exit Function
			If diffd = 2 Then DateTime = IIF(before,"ǰ�� ","���� ") & iiHour & ":" & iiMinute : Exit Function
			If diffd = 1 Then DateTime = IIF(before,"���� ","���� ") & iiHour & ":" & iiMinute : Exit Function
			DateTime = Int(diffs\3600) & "Сʱ" & dire : Exit Function
		End If
		If diffd < 7 Then DateTime = diffd & "��" & dire & " " & iiHour & ":" & iiMinute : Exit Function
		If diffd < 14 Then
			If diffw = 1 Then DateTime = IIF(before,"������","������") & iWeek(tWeek) & " " & iiHour & ":" & iiMinute : Exit Function
			If Not pastTime Then DateTime = diffd & "��" & dire : Exit Function
		End If
		If Not pastTime Then
			If diffd < 31 Then
				If diffm = 2 Then DateTime = "2����" & dire : Exit Function
				If diffm = 1 Then DateTime = IIF(before,"�ϸ���","�¸���") & iDay & "��" : Exit Function
				DateTime = diffw & "����" & dire : Exit Function
			End If
			If diffm < 36 Then
				If diffy = 3 Then DateTime = "3��" & dire : Exit Function
				If diffy = 2 Then DateTime = IIF(before,"ǰ��","����") & iMonth & "��" : Exit Function
				If diffy = 1 Then DateTime = IIF(before,"ȥ��","����") & iMonth & "��" : Exit Function
				DateTime = diffm & "����" & dire : Exit Function
			End If
			DateTime = diffy & "��" & dire : Exit Function
		Else
			iFormat = "yyyy-mm-dd hh:ii"
		End If
	End If
	iiWeek = Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday")
	iiiWeek = Array("Sun","Mon","Tue","Wed","Thu","Fri","Sat")
	iiiMonth = Array("Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec")
	iiiiMonth = Array("January","February","March","April","May","June","July","August","September","October","November","December")
	SpecialText = Array("y","m","d","h","i","s","w")
	SpecialTextRe = Array(Chr(0),Chr(1),Chr(2),Chr(3),Chr(4),Chr(5),Chr(6))
	For i = 0 To 6 : iFormat = Replace(iFormat,"\"&SpecialText(i), SpecialTextRe(i)) : Next
	t = Replace(iFormat,"yyyy", iiYear) : t = Replace(t, "yyy", iiYear)
	t = Replace(t, "yy", iYear) : t = Replace(t, "y", iiYear)
	t = Replace(t, "mmmm", iiiiMonth(iMonth-1)) : t = Replace(t, "mmm", iiiMonth(iMonth-1))
	t = Replace(t, "mm", iiMonth) : t = Replace(t, "m", iMonth)
	t = Replace(t, "dd", iiDay) : t = Replace(t, "d", iDay)
	t = Replace(t, "hh", iiHour) : t = Replace(t, "h", iHour)
	t = Replace(t, "ii", iiMinute) : t = Replace(t, "i", iMinute)
	t = Replace(t, "ss", iiSecond) : t = Replace(t, "s", iSecond)
	t = Replace(t, "www", iiiWeek(tWeek)) : t = Replace(t, "ww", iiWeek(tWeek))
	t = Replace(t, "w", iWeek(tWeek))
	For i = 0 To 6 : t = Replace(t, SpecialTextRe(i),SpecialText(i)) : Next
	DateTime = t
End Function
'��ȫ��ȡֵ
Function R(ByVal Str, ByVal RType)
	R = SafeData("R", Str, RType)
End Function
Function Ra(ByVal Str, ByVal RType)
	Ra = SafeData("Ra", Str, RType)
End Function
Function RF(ByVal Str, ByVal RType)
	RF = SafeData("RF", Str, RType)
End Function
Function RFa(ByVal Str, ByVal RType)
	RFa = SafeData("RFa", Str, RType)
End Function
Function RQ(ByVal Str, ByVal RType)
	RQ = SafeData("RQ", Str, RType)
End Function
Function RQa(ByVal Str, ByVal RType)
	RQa = SafeData("RQa", Str, RType)
End Function
Function RH(ByVal Str, ByVal RType)
	RH = SafeData("RH", Str, RType)
End Function
Function RHa(ByVal Str, ByVal RType)
	RHa = SafeData("RHa", Str, RType)
End Function
'��ȫ��ȡֵԭʼ����
Function SafeData(fn, ByVal Str, ByVal RType)
	Dim TempStr, fna
	Dim RDefault,RSplit
	Dim TempArr, i
	Select Case fn
		Case "R", "Ra" TempStr = Request(Str)
		Case "RF", "RFa" TempStr = Request.Form(Str)
		Case "RQ", "RQa" TempStr = Request.QueryString(Str)
		Case "RH", "RHa" TempStr = Request.QueryString()
		Case Else TempStr = Str
	End Select
	fna = IIF(fn = "Ra" or fn = "RFa" or fn = "RQa" or fn = "RHa",True,False)
	If fn = "RH" or fn = "RHa" Then
		If Not isNumeric(Str) Then SafeData = "" : Exit Function
		TempStr = Split(Split(TempStr,".")(0),"-")(Str)
	End If
	RSplit = ","
	If Instr(Cstr(RType),":")=2 Then
		RDefault = Mid(RType,3)
		If IsN(TempStr) Then TempStr = RDefault
		RType = Int(Left(RType,1))
		If RType = 2 Or RType = 3 Then RSplit = RDefault
	End If
	Select Case RType
		Case 0
			TempStr = Replace(TempStr,"'","''")
		Case 1
			TempStr = IsNumber(TempStr,IIF(fna,0,1))
		Case 2,3
			If Instr(TempStr,RSplit)>0 Then
				TempArr = split(TempStr,RSplit)
				TempStr = ""
				For i = 0 To Ubound(TempArr)
					If i <>0 Then TempStr = TempStr & RSplit
					If RType = 2 Then
						TempStr = TempStr & Replace(Trim(TempArr(i)),"'","''")
					Else
						TempArr(i) = IsNumber(Trim(TempArr(i)),IIF(fna,0,1))
						TempStr = TempStr & TempArr(i)
					End If
				Next
			Else
				TempStr = IIF(RType = 2,Replace(TempStr,"'","''"),IsNumber(TempStr,IIF(fna,0,1)))
			End If
	End Select
	SafeData = TempStr
End Function
'��ʱ��֤�Ƿ�Ϊ����
Private Function IsNumber(Str, iType)
	If Not IsN(Str) Then
		If not isNumeric(Str) Then
			If iType = 0 Then
				Alert "�������Ͳ���ȷ��"
			Else
				IsNumber = ""
			End If
		Else
			IsNumber = Str
		End if
	End If
End Function
'����ύ������Դ
Function CheckDataFrom()
	CheckDataFrom = True
	server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
	server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
	if mid(server_v1,8,len(server_v2))<>server_v2 then
		CheckDataFrom = False
	end if
end Function
Sub CheckDataFromA()
	If Not CheckDataFrom Then alert "��ֹ��վ���ⲿ�ύ���ݣ�"
end Sub
'��SQLע��ǿ���
Function CheckSql()
	Dim noSQLStr, noSQL, StrGet, StrPost, i, j
	noSQLStr = " and, or, insert, exec, select, delete, update, count, chr, mid, master, truncate, char, declare"
	noSQL = Split(noSQLStr,",")
	If Request.QueryString<>"" Then
		For Each StrGet In Request.QueryString
			For i = 0 To Ubound(noSQL)
				If Instr(Request.QueryString(StrGet),noSQL(i))>0 Then
					CheckSql = False
					exit Function
				End If
			Next
		Next
	End If
	If Request.Form<>"" Then
		For Each StrPost In Request.Form
			For j = 0 To Ubound(noSQL)
				If Instr(Request.Form(StrPost),noSQL(j))>0 Then
					CheckSql = False
					exit Function
				End If
			Next
		Next
	End If
	CheckSql = True
End Function
Sub CheckSqlA()
	If Not CheckSql Then alert "�����к��зǷ��ַ���"
End Sub
'��ȡ���ַ�����߲��ֲ���������Ŵ���
Function CutString(ByVal str, ByVal strlen)
	Dim l,t,c,i,d,f
	l = len(str) : t = 0 : d = "��" : f = Easp_Param(strlen)
	If Not isN(f(1)) Then : strlen = Int(f(0)) : d = f(1) : f = "" : End If
	For i = 1 to l
		c = Abs(Ascw(Mid(str,i,1)))
		t = IIF(c > 255,t + 2,t + 1)
		If t >= strlen Then
			CutString = Left(str,i) & d
			Exit For
		Else
			CutString = str
		End If
	Next
	CutString = Replace(CutString,vbCrLf,"")
End Function
'��ȡ��ǰ�ļ��ĵ�ַ
Function GetUrl(param)
	Dim script_name,url,dir
	Dim out,qitem,qtemp,i,hasQS,qstring
	script_name = Request.ServerVariables("SCRIPT_NAME")
	url = script_name
	dir  = Left(script_name,InstrRev(script_name,"/"))
	If isN(param) Then
		Dim ustart,uport
		With Request
			If .ServerVariables("HTTPS")="on" Then
				ustart = "https://"
				uport = IIF(Int(.ServerVariables("SERVER_PORT"))=443,"",":"&.ServerVariables("SERVER_PORT"))
			Else
				ustart = "http://"
				uport = IIF(Int(.ServerVariables("SERVER_PORT"))=80,"",":"&.ServerVariables("SERVER_PORT"))
			End If
			url = ustart & .ServerVariables("SERVER_NAME") & uport & script_name
			If Not IsN(.QueryString()) Then url = url & "?" & .QueryString()
			GetUrl = url : Exit Function
		End With
	End If
	If param = "0" Then : GetUrl = url : Exit Function
	If param = "2" Then : GetUrl = dir : Exit Function
	If InStr(param,":")>0 Then
		url = dir
		out = Mid(param,2)
		hasQS = IIF(isN(out),0,1)
	Else
		out = param : hasQS = 1
	End If
	If Not IsN(Request.QueryString()) Then
		If param="1" Or hasQS = 0 Then
			url = url & "?" & Request.QueryString()
		Else
			qtemp = "" : i = 0 : out = ","&out&","
			qstring = IIF(InStr(out,"-")>0,"Not InStr(out,"",-""&qitem&"","")>0","InStr(out,"",""&qitem&"","")>0")
			For Each qitem In Request.QueryString()
				If Eval(qstring) Then
					If i<>0 Then qtemp = qtemp & "&"
					qtemp = qtemp & qitem & "=" & Request.QueryString(qitem)
					i = i + 1
				End If
			Next
			If Not isN(qtemp) Then url = url & "?" & qtemp
		End If
	End If
	GetUrl = url
End Function
'��ȡ��ҳURL��ַ�������µ�URL����
Function GetUrlWith(ByVal urlParam, ByVal ParamAndValue)
	Dim u,s
	u = GetUrl(urlParam)
	s = GetUrl(0)
	If Left(urlParam,1)=":" Then s = Left(u,InstrRev(u,"/"))
	GetUrlWith = u & IIF(isN(Mid(u,len(s)+1)),"?","&") & paramAndValue
End Function
'��ȡ�û�IP��ַ
Function GetIP()
	Dim addr, x, y
	x = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	y = Request.ServerVariables("REMOTE_ADDR")
	addr = IIF(isN(x) or lCase(x)="unknown",y,x)
	If InStr(addr,".")=0 Then addr = "0.0.0.0"
	GetIP = addr
End Function
'����һ������ʱ������������������Сʱ��
Function DiffHour(theDate)
	DiffHour=DateDiff("h",Now,theDate)
End Function
'����һ������ʱ����������������������
Function DiffDay(theDate)
	DiffDay=DateDiff("d",Now,theDate)
End Function
'HTML���뺯��
Function HtmlEncode(ByVal str)
	If Not IsN(str) Then
		str = Replace(str, Chr(38), "&#38;")
		str = Replace(str, ">", "&gt;")
		str = Replace(str, "<", "&lt;")
		str = Replace(str, Chr(39), "&#39;")
		str = Replace(str, Chr(32), "&nbsp;")
		str = Replace(str, Chr(34), "&quot;")
		str = Replace(str, Chr(9), "&nbsp;&nbsp;&nbsp;&nbsp;")
		str = Replace(str, Chr(13), "")
		str = Replace(str, Chr(10), "<br />")
	End If
	HtmlEncode = str
End Function
'HTML���뺯��
Function HtmlDecode(ByVal str)
	If Not IsN(str) Then
		str = Replace(str, "<br/>", Chr(13)&Chr(10))
		str = Replace(str, "<br>", Chr(13)&Chr(10))
		str = Replace(str, "<br />", Chr(13)&Chr(10))
		str = Replace(str, "&nbsp;&nbsp;&nbsp;&nbsp;", Chr(9))
		str = Replace(str, "&amp;", Chr(38))
		str = Replace(str, "&#39;", Chr(39))
		str = Replace(str, "&apos;", Chr(39))
		str = Replace(str, "&nbsp;", Chr(32))
		str = Replace(str, "&quot;", Chr(34))
		str = Replace(str, "&gt;", ">")
		str = Replace(str, "&lt;", "<")
		str = Replace(str, "&#38;", Chr(38))
		HtmlDecode = str
	End If
End Function
'����HTML��ǩ
Function HtmlFilter(ByVal str)
	str = regReplace(str,"<[^>]+>|</[^>]+>","")
	str = Replace(str, ">", "&gt;")
	str = Replace(str, "<", "&lt;")
	HtmlFilter = str
End Function
'��ȷ������Ľű�ִ��ʱ��
Function GetScriptTime(StartTimer)
	GetScriptTime = FormatNumber((Timer()-StartTimer)*1000, 2, -1)
End Function
'ȡָ�����ȵ�����ַ���
Function RandStr(ByVal length)
	RandStr = Easp_RandStr(length)
End Function
'ȡһ�������
Function Rand(ByVal min, ByVal max)
    Rand = Easp_Rand(min,max)
End Function
'��ʽ������
Function toNumber(ByVal num, ByVal d)
	toNumber = FormatNumber(num,d,-1)
End Function
'������ת��Ϊ���Ҹ�ʽ
Function toPrice(ByVal num)
	toPrice = FormatCurrency(num,2,-1,0,-1)
End Function
'������ת��Ϊ�ٷֱȸ�ʽ
Function toPercent(ByVal num)
	toPercent = FormatPercent(num,2,-1)
End Function
'�رն����ͷ���Դ
Sub C(ByRef obj)
	On Error Resume Next
	obj.Close() : Set obj = Nothing
End Sub
'������ҳ����Ϣ
Sub noCache()
	Response.Buffer = True
	Response.Expires = 0
	Response.ExpiresAbsolute = Now() - 1
	Response.CacheControl = "no-cache"
	Response.AddHeader "Expires",Date()
	Response.AddHeader "Pragma","no-cache"
	Response.AddHeader "Cache-Control","private, no-cache, must-revalidate"
End Sub
'����һ��Cookiesֵ
Sub SetCookie(ByVal cooName, ByVal cooValue, ByVal cooCfg)
	Dim n,i,cExp,cDomain,cPath,cSecure
	If isArray(cooCfg) Then
		For i = 0 To Ubound(cooCfg)
			If Test(cooCfg(i),"date") Then
				cExp = cDate(cooCfg(i))
			ElseIf Test(cooCfg(i),"int") Then
				If cooCfg(i)<>0 Then cExp = Now()+Int(cooCfg(i))/60/24
			ElseIf Test(cooCfg(i),"domain") Then
				cDomain = cooCfg(i)
			ElseIf Instr(cooCfg(i),"/")>0 Then
				cPath = cooCfg(i)
			ElseIf cooCfg(i)="True" or cooCfg(i)="False" Then
				cSecure = cooCfg(i)
			End If
		Next
	Else
		If Test(cooCfg,"date") Then
			cExp = cDate(cooCfg)
		ElseIf Test(cooCfg,"int") Then
			If cooCfg<>0 Then cExp = Now()+Int(cooCfg)/60/24
		ElseIf Test(cooCfg,"domain") Then
			cDomain = cooCfg
		ElseIf Instr(cooCfg,"/")>0 Then
			cPath = cooCfg
		ElseIf cooCfg = "True" or cooCfg = "False" Then
			cSecure = cooCfg
		End If
	End If
	n = Easp_Param(cooName)
	If Not isN(cooValue) Then
		If isN(n(1)) Then
			Response.Cookies(n(0)) = cooValue
		Else
			Response.Cookies(n(0))(n(1)) = cooValue
		End If
	End If
	If Not isN(cExp) Then Response.Cookies(n(0)).Expires = cExp
	If Not isN(cDomain) Then Response.Cookies(n(0)).Domain = cDomain
	If Not isN(cPath) Then Response.Cookies(n(0)).Path = cPath
	If Not isN(cSecure) Then Response.Cookies(n(0)).Secure = cSecure
End Sub
'��ȡһ��Cookiesֵ
Function GetCookie(ByVal cooName)
	Dim n : n = Easp_Param(cooName)
	If Response.Cookies(n(0)).HasKeys And Not isN(n(1)) Then
		GetCookie = SafeData("",Request.Cookies(n(0))(n(1)),0)
	Else
		GetCookie = SafeData("",Request.Cookies(n(0)),0)
	End If
	If IsN(GetCookie) Then GetCookie = ""
End Function
'ɾ��һ��Cookiesֵ
Sub RemoveCookie(ByVal cooName)
	Dim n : n = Easp_Param(cooName)
	If Response.Cookies(n(0)).HasKeys And Not isN(n(1)) Then
		Response.Cookies(n(0))(n(1)) = Empty
	Else
		Response.Cookies(n(0)) = Empty
		Response.Cookies(n(0)).Expires = Now()
	End If
End Sub
'���û����¼
Sub SetApp(AppName,AppData)
	Application.Lock
	Application.Contents.Item(AppName) = AppData
	Application.UnLock
End Sub
'��ȡһ�������¼
Function GetApp(AppName)
	If IsN(GetApp) Then GetApp = "" : Exit Function
	GetApp = Application.Contents.Item(AppName)
End Function
'ɾ��һ�������¼
Sub RemoveApp(AppName)
	Application.Lock
	Application.Contents.Remove(AppName)
	Application.UnLock
End Sub
'��֤���֤����
Private Function isIDCard(ByVal str)
	Dim Ai, BirthDay, arrVerifyCode, Wi, i, AiPlusWi, modValue, strVerifyCode
	isIDCard = False
	If Len(str) <> 15 And Len(str) <> 18 Then Exit Function
	Ai = IIF(Len(str) = 18,Mid(str, 1, 17),Left(str, 6) & "19" & Mid(str, 7, 9))
	If Not IsNumeric(Ai) Then Exit Function
	If Not Test(Left(Ai,6),"^(1[1-5]|2[1-3]|3[1-7]|4[1-6]|5[0-4]|6[1-5]|8[12]|91)\d{2}[01238]\d{1}$") Then Exit Function
	BirthDay = Mid(Ai, 7, 4) & "-" & Mid(Ai, 11, 2) & "-" & Mid(Ai, 13, 2)
	If IsDate(BirthDay) Then
		If cDate(BirthDay) > Date() Or cDate(BirthDay) < cDate("1870-1-1") Then  Exit Function
	Else
		Exit Function
	End If
	arrVerifyCode = Split("1,0,x,9,8,7,6,5,4,3,2", ",")
	Wi = Split("7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2", ",")
	For i = 0 To 16
		AiPlusWi = AiPlusWi + CInt(Mid(Ai, i + 1, 1)) * Wi(i)
	Next
	modValue = AiPlusWi Mod 11
	strVerifyCode = arrVerifyCode(modValue)
	Ai = Ai & strVerifyCode
	If Len(str) = 18 And LCase(str) <> Ai Then Exit Function
	isIDCard = True
End Function
'���׵ķ���˼���
Function CheckForm(ByVal Str, ByVal Rule, ByVal Require, ByVal ErrMsg)
	Dim tmpMsg, Msg
	tmpMsg = Replace(ErrMsg,"\:",chr(0))
	Msg = IIF(Instr(tmpMsg,":")>0,Split(tmpMsg,":"),Array("����Ŀ����Ϊ��",tmpMsg))
	If Require = 1 And IsN(Str) Then
		If Instr(tmpMsg,":")>0 Then
			alert Replace(Msg(0),chr(0),":") : Exit Function
		Else
			alert Replace(tmpMsg,chr(0),":") : Exit Function
		End If
	End If
	If Not (Require = 0 And isN(Str)) Then
		If Left(Rule,1)=":" Then
			pass = False
			arrRule = Split(Mid(Rule,2),"|")
			For i = 0 To Ubound(arrRule)
				If Test(Str,arrRule(i)) Then pass = True : Exit For
			Next
			If Not pass Then Alert(Replace(Msg(1),chr(0),":")) : Exit Function
		Else
			If Not Test(Str,Rule) Then : Alert(Replace(Msg(1),chr(0),":")) : Exit Function
		End If
	End If
	CheckForm = Str
End Function
'����������֤���
Function Test(ByVal Str, ByVal Pattern)
	Dim Pa
	Select Case Lcase(Pattern)
		Case "date"		Test = IIF(isDate(Str),True,False) : Exit Function
		Case "idcard"	Test = IIF(isIDCard(Str),True,False) : Exit Function
		Case "english"	Pa = "^[A-Za-z]+$"
		Case "chinese"	Pa = "^[\u0391-\uFFE5]+$"
		Case "username"	Pa = "^[a-z]\w{2,19}$"
		Case "email"	Pa = "^\w+([-+.]\w+)*@\w+([-.]\w+)*\.\w+([-.]\w+)*$"
		Case "int"		Pa = "^[-\+]?\d+$"
		Case "number"	Pa = "^\d+$"
		Case "double"	Pa = "^[-\+]?\d+(\.\d+)?$"
		Case "price"	Pa = "^\d+(\.\d+)?$"
		Case "zip"		Pa = "^[1-9]\d{5}$"
		Case "qq"		Pa = "^[1-9]\d{4,9}$"
		Case "phone"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(\(0\d{2,3}\)|0\d{2,3}-)?[1-9]\d{6,7}(\-\d{1,4})?$"
		Case "mobile"	Pa = "^((\(\d{2,3}\))|(\d{3}\-))?(1[35][0-9]|189)\d{8}$"
		Case "url"		Pa = "^(http|https|ftp):\/\/[A-Za-z0-9]+\.[A-Za-z0-9]+[\/=\?%\-&_~`@[\]\':+!]*([^<>\""])*$"
		Case "domain"	Pa = "^[A-Za-z0-9\-]+\.([A-Za-z]{2,4}|[A-Za-z]{2,4}\.[A-Za-z]{2})$"
		Case "ip"		Pa = "^(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5]).(0|[1-9]\d?|[0-1]\d{2}|2[0-4]\d|25[0-5])$"
		Case Else Pa = Pattern
	End Select
	Test = Easp_Test(CStr(Str),Pa)
End Function
'�����滻
Function regReplace(ByVal Str, ByVal rule, Byval Result)
	regReplace = Easp_Replace(Str,rule,Result,0)
End Function
'�����滻����ģʽ
Function regReplaceM(ByVal Str, ByVal rule, Byval Result)
	regReplaceM = Easp_Replace(Str,rule,Result,1)
End Function
'����ƥ�䲶��
Function regMatch(ByVal Str, ByVal rule)
	Set regMatch =  Easp_Match(Str,rule)
End Function
'�������Ƿ�װ
Function isInstall(Byval Str)
	On Error Resume Next : Err.Clear()
	isInstall = False
	Dim obj : Set obj = Server.CreateObject(Str)
	If Err.Number = 0 Then isInstall = True
	Set obj = Nothing : Err.Clear()
End Function
'��̬�����ļ�
Sub Include(ByVal filePath)
	ExecuteGlobal GetIncCode(filePath,0)
End Sub
Function getInclude(ByVal filePath)
	ExecuteGlobal GetIncCode(filePath,1)
	getInclude = EasyAsp_s_html
End Function
'��ȡ�ļ�����
Private Function Read(ByVal filePath)
	Dim Fso, p, f, tmpStr
	p = filePath
	If Not (Mid(filePath,2,1)=":") Then p = Server.MapPath(filePath)
	Set Fso = Server.CreateObject(s_fsoName)
	If  Fso.FileExists(p) Then
		Set f = Fso.OpenTextFile(p)
		tmpStr = f.ReadAll
		f.Close()
		Set f = Nothing
	Else
		tmpStr = "�ļ�δ�ҵ�:" & filePath
	End If
	Set Fso = Nothing
	Read = tmpStr
End Function
'��ȡ�����ļ����ݣ����޼���
Private Function IncRead(ByVal filePath)
	Dim content, rule, inc, incFile, incStr
	content = Read(filePath)
	If isN(content) Then Exit Function
	content = regReplace(content,"<% *?@.*?%"&">","")
	content = regReplace(content,"(<%[^>]+?)(option +?explicit)([^>]*?%"&">)","$1'$2$3")
	rule = "<!-- *?#include +?(file|virtual) *?= *?""??([^"":?*\f\n\r\t\v]+?)""?? *?-->"
	If Easp_Test(content,rule) Then
		Set inc = regMatch(content,rule)
		For Each Match In inc
			If LCase(Match.SubMatches(0))="virtual" Then
				incFile = Match.SubMatches(1)
			Else
				incFile = Mid(filePath,1,InstrRev(filePath,IIF(Instr(filePath,":")>0,"\","/"))) & Match.SubMatches(1)
			End If
			incStr = IncRead(incFile)
			content = Replace(content,Match,incStr)
		Next
		Set inc = Nothing
	End If
	IncRead = content
End Function
'�������ļ�ת��ΪASP����
Private Function GetIncCode(ByVal filePath, ByVal getHtml)
	Dim content,tmpStr,code,tmpCode,s_code,st,en
	content = IncRead(filePath)
	code = "" : st = 1 : en = Instr(content,"<%") + 2
	s_code = IIF(getHtml=1,"EasyAsp_s_html = EasyAsp_s_html & ","Response.Write ")
	While en > st + 1
		tmpStr = Mid(content,st,en-st-2)
		st = Instr(en,content,"%"&">") + 2
		If Not isN(tmpStr) Then
			tmpStr = Replace(tmpStr,"""","""""")
			tmpStr = Replace(tmpStr,vbCrLf&vbCrLf,vbCrLf)
			tmpStr = Replace(tmpStr,vbCrLf,"""&vbCrLf&""")
			code = code & s_code & """" & tmpStr & """" & vbCrLf
		End If
		tmpStr = Mid(content,en,st-en-2)
		tmpCode = regReplace(tmpStr,"^\s*=\s*",s_code) & vbCrLf
		If getHtml = 1 Then
			tmpCode = regReplaceM(tmpCode,"^(\s*)response\.write","$1" & s_code) & vbCrLf
		End If
		code = code & Replace(tmpCode,vbCrLf&vbCrLf,vbCrLf)
		en = Instr(st,content,"<%") + 2
	Wend
	tmpStr = Mid(content,st)
	If Not isN(tmpStr) Then
		tmpStr = Replace(tmpStr,"""","""""")
		tmpStr = Replace(tmpStr,vbCrLf&vbCrLf,vbCrLf)
		tmpStr = Replace(tmpStr,vbcrlf,"""&vbCrLf&""")
		code = code & s_code & """" & tmpStr & """" & vbCrLf
	End If
	If getHtml = 1 Then code = "EasyAsp_s_html = """" " & vbCrLf & code
	GetIncCode = Replace(code,vbCrLf&vbCrLf,vbCrLf)
End Function
End Class
'***** ���ݿ������ *****
Class EasyAsp_db
	Private idbConn, idbType, idebug, idbErr, iQueryType
	Private iPageParam, iPageIndex, iPageSize, iPageSpName, iPageCount, iRecordCount, iPageDic

	Private Sub Class_Initialize()
		'On Error Resume Next
		idbType = ""
		idebug = False
		idbErr = ""
		iQueryType = 0
		If TypeName(Conn) = "Connection" Then
			Set idbConn = Conn : idbType = GetDataType(Conn)
		End If
		iPageParam = "page"							'Ĭ��ҳ��URL������
		iPageSize = 20								'Ĭ��ÿҳ��¼��
		iPageSpName = "easp_sp_pager"				'Ĭ��MSSQL��ҳ�洢������
		Set iPageDic = Server.CreateObject("Scripting.Dictionary")
		'Ĭ�Ϸ�ҳ������ʽ
		iPageDic("default_html") = "<div class=""pager"">{first}{prev}{liststart}{list}{listend}{next}{last} ��ת��{jump}ҳ</div>"
		iPageDic("default_config") = ""
	End Sub
	Private Sub Class_Terminate()
		If TypeName(idbConn) = "Connection" Then
			If idbConn.State = 1 Then idbConn.Close()
			Set idbConn = Nothing
		End If
		Set iPageDic = Nothing
	End Sub
	'���ԣ��������ݿ�����
	Public Property Let dbConn(ByVal pdbConn)
		If TypeName(pdbConn) = "Connection" Then
			Set idbConn = pdbConn
			idbType = GetDataType(pdbConn)
		Else
			ErrMsg "��Ч�����ݿ�����", Err.Description
		End If
	End Property
	Public Property Get dbConn()
		Set dbConn = idbConn
	End Property
	'���ԣ���ǰ���ݿ�����
	Public Property Get DatabaseType()
		DatabaseType = idbType
	End Property
	'���ԣ������Ƿ�������ģʽ
	Public Property Let Debug(ByVal bool)
		idebug = bool
	End Property
	Public Property Get Debug()
		Debug = idebug
	End Property
	'���ԣ����ش�����Ϣ
	Public Property Get dbErr()
		dbErr = idbErr
	End Property
	'���ԣ����û�ȡ��¼���ķ�ʽ
	Public Property Let QueryType(ByVal str)
		str = Lcase(str)
		If str = "1" or str = "command" Then
			iQueryType = 1
		Else
			iQueryType = 0
		End If
	End Property
	'���ԣ����÷�ҳ����
	Public Property Let PageSize(ByVal num)
		iPageSize = num
	End Property
	'���ԣ����ط�ҳ����
	Public Property Get PageSize()
		PageSize = iPageSize
	End Property
	'���ԣ�������ҳ��
	Public Property Get PageCount()
		PageCount = iPageCount
	End Property
	'���ԣ����ص�ǰҳ��
	Public Property Get PageIndex()
		PageIndex = Easp_IIF(Easp_isN(iPageIndex),GetCurrentPage,iPageIndex)
	End Property
	'���ԣ������ܼ�¼��
	Public Property Get PageRecordCount()
		PageRecordCount = iRecordCount
	End Property
	'���ԣ����û�ȡ��ҳ����
	Public Property Let PageParam(ByVal str)
		iPageParam = str
	End Property
	'���ԣ����÷�ҳ�洢������
	Public Property Let PageSpName(ByVal str)
		iPageSpName = str
	End Property
	Private Sub ErrMsg(e,d)
		idbErr = "<div id=""easp_db_err"">" & e
		If d<>"" Then idbErr = idbErr & "<br/>������Ϣ��" & d
		idbErr = idbErr & "</div>"
		If idebug Then
			Response.Write idbErr
			Response.End()
		End If
	End Sub
	'�������ݿ������ַ���
	Public Function OpenConn(ByVal dbType, ByVal strDB, ByVal strServer)
		Dim TempStr, objConn, s, u, p, port
		s = "" : u = "" : p = "" : port = ""
		If Instr(strServer,"@")>0 Then
			s = Trim(Mid(strServer,InstrRev(strServer,"@")+1))
			u = Trim(Left(strServer,InstrRev(strServer,"@")-1))
			If Instr(s,":")>0 Then : port = Trim(Mid(s,Instr(s,":")+1)) : s = Trim(Left(s,Instr(s,":")-1))
			If Instr(u,":")>0 Then : p = Trim(Mid(u,Instr(u,":")+1)) : u = Trim(Left(u,Instr(u,":")-1))
		Else
			If Instr(strServer,":")>0 Then
				u = Trim(Left(strServer,Instr(strServer,":")-1))
				p = Trim(Mid(strServer,Instr(strServer,":")+1))
			Else
				p = Trim(strServer)
			End If
		End If
		idbType = UCase(Cstr(dbType))
		Select Case idbType
			Case "0","MSSQL"
				TempStr = "driver={sql server};server="&s&";uid="&u&";pwd="&p&";database="&strDB
			Case "1","ACCESS"
				Dim tDb : If Instr(strDB,":")>0 Then : tDb = strDB : Else : tDb = Server.MapPath(strDB) : End If
				TempStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&tDb&";Jet OLEDB:Database Password="&p&";"
			Case "2","MYSQL"
				If port = "" Then port = "3306"
				TempStr = "Driver={mySQL};Server="&s&";Port="&port&";Option=131072;Stmt=;Database="&strDB&";Uid="&u&";Pwd="&p&";"
			Case "3","ORACLE"
				TempStr = "Driver={Microsoft ODBC for Oracle};Server="&s&";Uid="&u&";Pwd="&p&";"
		End Select
		Set OpenConn = CreatConn(TempStr)
	End Function
	'�������ݿ����Ӷ���
	Public Function CreatConn(ByVal ConnStr)
		On Error Resume Next
		Dim objConn : Set objConn = Server.CreateObject("ADODB.Connection")
		objConn.Open ConnStr
		If Err.number <> 0 Then
			ErrMsg "���ݿ�����������Ӵ����������ݿ����ӡ�", Err.Description
			objConn.Close
			Set objConn = Nothing
		End If
		Set CreatConn = objConn
	End Function
	Private Function GetDataType(ByVal connObj)
		Dim str,i : str = UCase(connObj.Provider)
		Dim MSSQL, ACCESS, MYSQL, ORACLE
		MSSQL = Split("SQLNCLI10, SQLXMLOLEDB, SQLNCLI, SQLOLEDB, MSDASQL",", ")
		ACCESS = Split("MICROSOFT.ACE.OLEDB.12.0, MICROSOFT.JET.OLEDB.4.0",", ")
		MYSQL = "MYSQLPROV"
		ORACLE = Split("MSDAORA, OLEDB.ORACLE",", ")
		For i = 0 To Ubound(MSSQL)
			If Instr(str,MSSQL(i))>0 Then
				GetDataType = "MSSQL" : Exit Function
			End If
		Next
		For i = 0 To Ubound(ACCESS)
			If Instr(str,ACCESS(i))>0 Then
				GetDataType = "ACCESS" : Exit Function
			End If
		Next
		If Instr(str,MYSQL)>0 Then
			GetDataType = "MYSQL" : Exit Function
		End If
		For i = 0 To Ubound(ORACLE)
			If Instr(str,ORACLE(i))>0 Then
				GetDataType = "ORACLE" : Exit Function
			End If
		Next
	End Function
	'�Զ���ȡΨһ���кţ��Զ���ţ�
	Public Function AutoID(ByVal TableName)
		On Error Resume Next
		Dim rs, tmp, fID, tmpID : fID = "" : tmpID = 0
		tmp = Easp_Param(TableName)
		If Not Easp_isN(tmp(1)) Then : TableName = tmp(0) : fID = tmp(1) : tmp = "" : End If
		Set rs = GRS("Select " & Easp_IIF(fID<>"", "Max("&fID&")", "Top 1 *") & " From ["&TableName&"]")
		If rs.eof Then
			AutoID = 1 : Exit Function
		Else
			If fID<>"" Then
				If Easp_isN(rs.Fields.Item(0).Value) Then AutoID = 1 : Exit Function
				AutoID = rs.Fields.Item(0).Value + 1 : Exit Function
			Else
				Dim newRs
				Set newRs = GRS("Select Max("&rs.Fields.Item(0).Name&") From ["&TableName&"]")
				tmpID = newRS.Fields.Item(0).Value + 1
				newRs.Close() : Set newRs = Nothing
			End If
		End If
		If Err.number <> 0 Then ErrMsg "��Ч�Ĳ�ѯ�������޷���ȡ�µ�ID�ţ�", Err.Description
		rs.Close() : Set rs = Nothing
		AutoID = tmpID
	End Function
	'ȡ�÷��������ļ�¼�б�
	Public Function GetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GetRecord = GRS(wGetRecord(TableName,Condition,OrderField))
	End Function
	
	Public Function wGetRecord(ByVal TableName,ByVal Condition,ByVal OrderField)
		Dim strSelect, FieldsList, ShowN, o, p
		FieldsList = "" : ShowN = 0
		o = Easp_Param(TableName)
		If Not Easp_isN(o(1)) Then
			TableName = Trim(o(0)) : FieldsList = Trim(o(1)) : o = ""
			p = Easp_Param(FieldsList)
			If Not Easp_isN(p(1)) Then
				FieldsList = Trim(p(0)) : ShowN = Int(Trim(p(1))) : p = ""
			Else
				If isNumeric(FieldsList) Then ShowN = Int(FieldsList) : FieldsList = ""
			End If
		End If
		strSelect = "Select "
		If ShowN > 0 Then strSelect = strSelect & "Top " & ShowN & " "
		strSelect = strSelect & Easp_IIF(FieldsList <> "", FieldsList, "* ")
		strSelect = strSelect & " From [" & TableName & "]"
		If isArray(Condition) Then
			strSelect = strSelect & " Where " & ValueToSql(TableName,Condition,1)
		Else
			If Condition <> "" Then strSelect = strSelect & " Where " & Condition
		End If
		If OrderField <> "" Then strSelect = strSelect & " Order By " & OrderField
		wGetRecord = strSelect
	End Function
	Public Function GR(ByVal TableName,ByVal Condition,ByVal OrderField)
		Set GR = GetRecord(TableName, Condition, OrderField)
	End Function
	Public Function wGR(ByVal TableName,ByVal Condition,ByVal OrderField)
		wGR = wGetRecord(TableName, Condition, OrderField)
	End Function
	'����sql��䷵�ؼ�¼��
	Public Function GetRecordBySQL(ByVal str)
		On Error Resume Next
		If iQueryType = 1 Then
			Dim cmd : Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = idbConn
				.CommandText = str
				Set GetRecordBySQL = .Execute
			End With
			Set cmd = Nothing
		Else
			Dim rs : Set rs = Server.CreateObject("Adodb.Recordset")
			With rs
				.ActiveConnection = idbConn
				.CursorType = 1
				.LockType = 1
				.Source = str
				.Open
			End With
			Set GetRecordBySQL = rs
		End If
		If Err.number <> 0 Then ErrMsg "��Ч�Ĳ�ѯ�������޷���ȡ��¼����", Err.Description & "<br/>SQL��" & str
		Err.Clear
	End Function
	Public Function GRS(ByVal strSelect)
		Set GRS = GetRecordBySQL(strSelect)
	End Function
	'���ݼ�¼������Json��ʽ����
	Public Function Json(ByVal jRs, ByVal jName)
		On Error Resume Next
		Dim tmpStr, rs, fi, i, j, o, isE,tName,tValue : i = 0
		isE = False
		o = Easp_Param(jName)
		If Not Easp_isN(o(1)) Then
			jName = o(0)
			isE = True
		End If
		Set rs = jRs
		tmpStr = "{ """&jName&""" : ["
		rs.MoveFirst()
		If Not rs.bof And Not rs.eof Then
			While Not rs.Eof
				j = 0 : If i<>0 Then tmpStr = tmpStr & ", "
				tmpStr = tmpStr & "{"
				For Each fi In rs.Fields
					If j<>0 Then tmpStr = tmpStr & ", "
					tName = fi.Name : tValue = fi.Value
					If isE Then
						tmpStr = tmpStr & """" & Easp_Escape(tName) & """:""" & Easp_Escape(Easp_jsEncode(tValue)) & """"
					Else
						tmpStr = tmpStr & """" & tName & """:""" & Easp_jsEncode(tValue) & """"
					End If
					j = j + 1
				Next
				tmpStr = tmpStr & "}"
				i = i + 1 : rs.MoveNext()
			Wend
		End If
		tmpStr = tmpStr & "]}"
		If Err.number <> 0 Then ErrMsg "����Json��ʽ�������", Err.Description
		rs.Close() : Set rs = Nothing
		Json = tmpStr
	End Function
	'����ָ�����ȵĲ��ظ����ַ���
	Public Function RandStr(length,TableField)
		On Error Resume Next
		Dim tb, fi, tmpStr, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpStr = Easp_RandStr(length)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",fi&"='"&tmpStr&"'","")
			If Not rs.Bof And Not rs.Eof Then
				tmpStr = Easp_RandStr(length)
			Else
				RandStr = tmpStr
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then ErrMsg "���ɲ��ظ�������ַ�������", Err.Description
	End Function
	'����һ�����ظ��������
	Public Function Rand(min,max,TableField)
		On Error Resume Next
		Dim tb, fi, tmpInt, rs
		tb = Easp_Param(TableField)(0)
		fi = Easp_Param(TableField)(1)
		tmpInt = Easp_Rand(min,max)
		Do While (True)
			Set rs = GR(tb&":"&fi&":1",Array(fi&":"&tmpInt),"")
			If Not rs.Bof And Not rs.Eof Then
				tmpInt = Easp_Rand(min,max)
			Else
				Rand = tmpInt
				Exit Do
			End If
			C(rs)
		Loop
		If Err.number <> 0 Then ErrMsg "���ɲ��ظ������������", Err.Description
	End Function
	'ȡ��ĳһָ����¼����ϸ����
	Public Function GetRecordDetail(ByVal TableName,ByVal Condition)
		Dim strSelect
		strSelect = "Select * From [" & TableName & "] Where " & ValueToSql(TableName,Condition,1)
		Set GetRecordDetail = GRS(strSelect)
	End Function
	Public Function GRD(ByVal TableName,ByVal Condition)
		Set GRD = GetRecordDetail(TableName, Condition)
	End Function
	'ȡָ�������������¼
	Public Function GetRandRecord(ByVal TableName,ByVal Condition)
		Dim sql,o,p,fi,IdField,showN,where
		o = Easp_Param(TableName)
		If Not Easp_isN(o(1)) Then
			TableName = o(0)
			p = Easp_Param(o(1))
			If Easp_isN(p(1)) Then
				ErrMsg "��ȡ�����¼ʧ�ܣ�", "������Ҫȡ�ļ�¼����"
				Exit Function
			Else
				fi = p(0) : showN = p(1)
				If Instr(fi,",")>0 Then
					IdField = Trim(Left(fi,Instr(fi,",")-1))
				Else
					IdField = fi : fi = "*"
				End If
			End If
		Else
			ErrMsg "��ȡ�����¼ʧ�ܣ�", "���ڱ���������:ID�ֶε�����"
			Exit Function
		End If
		Condition = Easp_IIF(Easp_isN(Condition),""," Where " & ValueToSql(TableName,Condition,1))
		sql = "Select Top " & showN & " " & fi & " From ["&TableName&"]" & Condition
		Select Case idbType
			Case "ACCESS" : Randomize
				sql = sql & " Order By Rnd(-(" & IdField & "+" & Rnd() & "))"
			Case "MSSQL"
				sql = sql & " Order By newid()"
			Case "MYSQL"
				sql = "Select " & fi & " From ["&TableName&"]" & Condition & " Order By rand() limit " & showN
			Case "ORACLE"
				sql = "Select " & fi & " From (Select " & fi & " From ["&TableName&"] Order By dbms_random.value) " & Easp_IIF(Easp_isN(Condition),"Where",Condition & " And") & " rownum < " & Int(showN)+1
		End Select
		Set GetRandRecord = GRS(sql)
	End Function
	Public Function GRR(ByVal TableName,ByVal Condition)
		Set GRR = GetRandRecord(TableName,Condition)
	End Function
	'���һ���µļ�¼
	Public Function AddRecord(ByVal TableName,ByVal ValueList)
		On Error Resume Next
		Dim o : o = Easp_Param(TableName) : If Not Easp_isN(o(1)) Then TableName = o(0)
		DoExecute wAddRecord(TableName,ValueList)
		If Err.number <> 0 Then
			ErrMsg "�����ݿ���Ӽ�¼����", Err.Description
			AddRecord = 0
			Exit Function
		End If
		If Not Easp_isN(o(1)) Then
			AddRecord = AutoID(o(0)&":"&o(1))-1
		Else
			AddRecord = 1
		End If
	End Function
	Public Function wAddRecord(ByVal TableName,ByVal ValueList)
		Dim TempSQL, TempFiled, TempValue, o
		o = Easp_Param(TableName) : If Not Easp_isN(o(1)) Then TableName = o(0)
		TempFiled = ValueToSql(TableName,ValueList,2)
		TempValue = ValueToSql(TableName,ValueList,3)
		TempSQL = "Insert Into [" & TableName & "] (" & TempFiled & ") Values (" & TempValue & ")"
		wAddRecord = TempSQL
	End Function
	Public Function AR(ByVal TableName,ByVal ValueList)
		AR = AddRecord(TableName,ValueList)
	End Function
	Public Function wAR(ByVal TableName,ByVal ValueList)
		wAR = wAddRecord(TableName,ValueList)
	End Function
	'�޸�ĳһ��¼
	Public Function UpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		On Error Resume Next
		DoExecute wUpdateRecord(TableName,Condition,ValueList)
		If Err.number <> 0 Then
			ErrMsg "�������ݿ��¼����", Err.Description
			UpdateRecord = 0
			Exit Function
		End If
		UpdateRecord = 1
	End Function
	Public Function wUpdateRecord(ByVal TableName,ByVal Condition,ByVal ValueList)
		Dim TmpSQL
		TmpSQL = "Update ["&TableName&"] Set "
		TmpSQL = TmpSQL & ValueToSql(TableName,ValueList,0)
		TmpSQL = TmpSQL & " Where " & ValueToSql(TableName,Condition,1)
		wUpdateRecord = TmpSQL
	End Function
	Public Function UR(ByVal TableName,ByVal Condition,ByVal ValueList)
		UR = UpdateRecord(TableName, Condition, ValueList)
	End Function
	Public Function wUR(ByVal TableName,ByVal Condition,ByVal ValueList)
		wUR = wUpdateRecord(TableName, Condition, ValueList)
	End Function
	'ɾ��ָ���ļ�¼
	Public Function DeleteRecord(ByVal TableName,ByVal Condition)
		On Error Resume Next
		DoExecute wDeleteRecord(TableName,Condition)
		If Err.number <> 0 Then
			ErrMsg "�����ݿ�ɾ�����ݳ���", Err.Description
			DeleteRecord = 0
			Exit Function
		End If
		DeleteRecord = 1
	End Function
	Public Function wDeleteRecord(ByVal TableName,ByVal Condition)
		Dim IDFieldName, IDValues, Sql, p : IDFieldName = "" : IDValues = ""
		If Not isArray(Condition) Then
			p = Easp_Param(Condition)
			If Not Easp_isN(p(1)) Then
				IDFieldName = p(0)
				If Instr(IDFieldName," ")=0 Then
					IDValues = p(1)
				Else
					IDFieldName = ""
				End If
			End If
		End If
		Sql = "Delete From ["&TableName&"] Where " & Easp_IIF(IDFieldName="", ValueToSql(TableName,Condition,1), "["&IDFieldName&"] In (" & IDValues & ")")
		wDeleteRecord = Sql
	End Function
	Public Function DR(ByVal TableName,ByVal Condition)
		DR = DeleteRecord(TableName, Condition)
	End Function
	Public Function wDR(ByVal TableName,ByVal Condition)
		wDR = wDeleteRecord(TableName, Condition)
	End Function
	'��ĳһ���У�����һ��������ȡһ����¼�������ֶε�ֵ
	Public Function ReadTable(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		On Error Resume Next
		Dim rs,Sql,arrTemp,arrStr,TempStr,i
		TempStr = "" : arrStr = ""
		Sql = "Select "&GetFieldNames&" From ["&TableName&"] Where " & ValueToSql(TableName,Condition,1)
		Set rs = GRS(Sql)
		If Not rs.Eof Then
			If Instr(GetFieldNames,",") > 0 Then
				arrTemp = Split(GetFieldNames,",")
				For i = 0 To Ubound(arrTemp)
					If i<>0 Then arrStr = arrStr & Chr(0)
					arrStr = arrStr & rs.Fields.Item(i).Value
				Next
				TempStr = Split(arrStr,Chr(0))
			Else
				TempStr = rs.Fields.Item(0).Value
			End If
		End If
		If Err.number <> 0 Then ErrMsg "�����ݿ��ȡ���ݳ���", Err.Description
		rs.close() : Set rs = Nothing : Err.Clear
		ReadTable = TempStr
	End Function
	Public Function RT(ByVal TableName,ByVal Condition,ByVal GetFieldNames)
		RT = ReadTable(TableName, Condition, GetFieldNames)
	End Function
	'���ô洢����
	Public Function doSP(ByVal spName, ByVal spParam)
		On Error Resume Next
		Dim p, spType, cmd, outParam, i, NewRS : spType = ""
		If Not idbType="0" And Not idbType="MSSQL" Then
			MsgErr "��֧�ִ�MS SQL Server���ݿ���ô洢���̣�",""
			Exit Function
		End If
		p = Easp_Param(spName)
		If Not Easp_isN(p(1)) Then : spType = UCase(Trim(p(1))) : spName = Trim(p(0)) : p = "" : End If
		Set cmd = Server.CreateObject("ADODB.Command")
			With cmd
				.ActiveConnection = idbConn
				.CommandText = spName
				.CommandType = 4
				.Prepared = true
				.Parameters.append .CreateParameter("return",3,4)
				outParam = "return"
				If Not IsArray(spParam) Then
					If spParam<>"" Then
						spParam = Easp_IIF(Instr(spParam,",")>0, spParam = Split(spParam,","), Array(spParam))
					End If
				End If
				If IsArray(spParam) Then
					For i = 0 To Ubound(spParam)
						Dim pName, pValue
						If (spType = "1" or spType = "OUT" or spType = "3" or spType = "ALL") And Instr(spParam(i),"@@")=1 Then
							.Parameters.append .CreateParameter(spParam(i),200,2,8000)
							outParam = outParam & "," & spParam(i)
						Else
							If Instr(spParam(i),"@")=1 And Instr(spParam(i),":")>2 Then
								pName = Left(spParam(i),Instr(spParam(i),":")-1)
								outParam = outParam & "," & pName
								pValue = Mid(spParam(i),Instr(spParam(i),":")+1)
								If pValue = "" Then pValue = NULL
								.Parameters.append .CreateParameter(pName,200,1,8000,pValue)
							Else
								.Parameters.append .CreateParameter("@param"&(i+1),200,1,8000,spParam(i))
								outParam = outParam & "," & "@param"&(i+1)
							End If
						End If
					Next
				End If
			End With
			outParam = Easp_IIF(Instr(outParam,",")>0, Split(outParam,","), Array(outParam))
			If spType = "1" or spType = "OUT" Then
				cmd.Execute : doSP = cmd
			ElseIf spType = "2" or spType = "RS" Then
				Set doSP = cmd.Execute
			ElseIf spType = "3" or spType = "ALL" Then
				Dim NewOut,pa : Set NewOut = Server.CreateObject("Scripting.Dictionary")
				Set NewRS = cmd.Execute : NewRS.close
				For i = 0 To Ubound(outParam)
					NewOut(Trim(outParam(i))) = cmd(i)
				Next
				NewRs.open : doSP = Array(NewRS,NewOut)
				Set NewOut = Nothing
			Else
				cmd.Execute : doSP = cmd(0)
			End If
		If Err.number <> 0 Then ErrMsg "���ô洢���̳���", Err.Description
		Set cmd = Nothing
		Err.Clear
	End Function
	'�ͷż�¼������
	Public Function C(ByRef ObjRs)
		On Error Resume Next
		ObjRs.close()
		Set ObjRs = Nothing
	End Function
	'ִ��ָ����SQL���,�ɷ��ؼ�¼��
	Public Function Exec(ByVal str)
		On Error Resume Next
		If Lcase(Left(str,6)) = "select" Then
			Dim i : i = iQueryType
			iQueryType = 1
			Set Exec = GRS(str)
			iQueryType = i
		Else
			Exec = 1 : DoExecute(str)
			If Err.number <> 0 Then Exec = 0
		End If
		If Err.number <> 0 Then
			ErrMsg "ִ��SQL������", Err.Description
		End If
		Err.Clear
	End Function
	
	Private Function ValueToSql(ByVal TableName, ByVal ValueList, ByVal sType)
		On Error Resume Next
		Dim StrTemp : StrTemp = ValueList
		If IsArray(ValueList) Then
			StrTemp = ""
			Dim rsTemp, CurrentField, CurrentValue, i
			Set rsTemp = GRS("Select * From [" & TableName & "] Where 1 = -1")
			For i = 0 to Ubound(ValueList)
				CurrentField = Easp_Param(ValueList(i))(0)
				CurrentValue = Easp_Param(ValueList(i))(1)
				If i <> 0 Then StrTemp = StrTemp & Easp_IIF(sType=1, " And ", ", ")
				If sType = 2 Then
					StrTemp = StrTemp & "[" & CurrentField & "]"
				Else
					Select Case rsTemp.Fields(CurrentField).Type
						Case 7,8,129,130,133,134,135,200,201,202,203
							StrTemp = StrTemp & Easp_IIF(sType=3, "'"&CurrentValue&"'", "[" & CurrentField & "] = '"&CurrentValue&"'")
						Case 11
							Dim tmpTF
							tmpTF = Easp_IIF(UCase(cstr(Trim(CurrentValue)))="TRUE", Easp_IIF(idbType="1" Or idbType="ACCESS","True","1"), Easp_IIF(idbType="1" Or idbType="ACCESS","False","0"))
							StrTemp = StrTemp & Easp_IIF(sType = 3, tmpTF, "[" & CurrentField & "] = " & tmpTF)
						Case Else
							StrTemp = StrTemp & Easp_IIF(sType = 3, CurrentValue, "[" & CurrentField & "] = " & CurrentValue)
					End Select
				End If
			Next
			If Err.number <> 0 Then ErrMsg "����SQL������", Err.Description
			rsTemp.Close() : Set rsTemp = Nothing : Err.Clear
		End If
		ValueToSql = StrTemp
	End Function
	Private Function DoExecute(ByVal sql)
		Dim ExecuteCmd : Set ExecuteCmd = Server.CreateObject("ADODB.Command")
		With ExecuteCmd
			.ActiveConnection = idbConn
			.CommandText = sql
			.Execute
		End With
		Set ExecuteCmd = Nothing
	End Function
	'�����Ƿ�ҳ���򲿷�
	'��ȡ��ҳ��ļ�¼��
	Public Function GetPageRecord(ByVal PageSetup, ByVal Condition)
		On Error Resume Next
		Dim pType,spResult,rs,o,p,Sql,n,i,spReturn
		o = Easp_Param(Cstr(PageSetup))
		pType = o(0)
		If Not Easp_isN(o(1)) Then
			p = Easp_Param(o(1))
			If Not Easp_isN(p(1)) Then
				iPageParam = Lcase(p(0))
				iPageSize = Int(p(1))
			Else
				If isNumeric(o(1)) Then
					iPageSize = Int(o(1))
				Else
					iPageParam = Lcase(o(1))
				End If
			End If
		End If
		iPageIndex = GetCurrentPage()
		Select Case Lcase(pType)
			Case "array","0"
				If isArray(Condition) Then
					Dim Table,Fi,Where
					o = Easp_Param(Condition(0))
					If Not Easp_isN(o(1)) Then
						Table = o(0) : Fi = o(1)
					Else
						Table = Condition(0) : Fi = "*"
					End If
					If isArray(Condition(1)) Then
						Where = ValueToSql(Table,Condition(1),1)
					Else
						Where = Condition(1)
					End If
					iRecordCount = Int(RT(Table, Easp_IIF(Easp_isN(Where),"1=1",Where), "Count(0)"))
					n = iRecordCount / iPageSize
					iPageCount = Easp_IIF(n=Int(n), n, Int(n)+1)
					iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
					If idbType = "1" or idbType = "ACCESS" Then
						Set rs = GR(Table&":"&Fi,Where,Condition(2))
						rs.PageSize = iPageSize
						If iRecordCount>0 Then rs.AbsolutePage = iPageIndex
						Set GetPageRecord = rs : Exit Function
					ElseIf idbType = "2" or idbType = "MYSQL" Then
						Sql = "Select "& fi & " From [" & Table & "]"
						If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
						If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
						Sql = Sql & " Limit " & iPageSize*(iPageIndex-1) & ", " & iPageSize
					Else
						If Ubound(Condition)<>3 Then ErrMsg "��ȡ��ҳ���ݳ���", "���������4��Ԫ�أ������ṩ���ݿ�����������"
						Sql = "Select Top " & iPageSize & " " & fi
						Sql = Sql & " From [" & Table & "]"
						If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
						If iPageIndex > 1 Then
							Sql = Sql & " " & Easp_IIF(Easp_isN(Where), "Where", "And") & " " & Condition(3) & " Not In ("
							Sql = Sql & "Select Top " & iPageSize * (iPageIndex-1) & " " & Condition(3) & " From [" & Table & "]"
							If Not Easp_isN(Where) Then Sql = Sql & " Where " & Where
							If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
							Sql = Sql & ") "
						End If
						If Not Easp_isN(Condition(2)) Then Sql = Sql & " Order By " & Condition(2)
					End If
					Set GetPageRecord = GRS(Sql)
				Else
					ErrMsg "��ȡ��ҳ���ݳ���", "ʹ������������ȡ��ҳ����ʱ������������Ϊ���飡"
				End If
			Case "sql","1" Set rs = GRS(Condition)
			Case "rs","2" Set rs = Condition
			Case Else
				If isArray(Condition) Then
					If pType = "" Then pType = iPageSpName
					Select Case pType
						Case "easp_sp_pager"	'ʹ���Դ���ҳ�洢���̷�ҳ
							If Ubound(Condition)<>5 Then ErrMsg "��ȡ��ҳ���ݳ���", "ʹ���Դ���ҳ�洢����ʱ���������������Ϊ6��Ԫ�أ�"
							spResult = doSP("easp_sp_pager:3",Array("@TableName:"&Condition(0),"@FieldList:"&Condition(1),"@Where:"&Condition(2),"@Order:"&Condition(3),"@PrimaryKey:"&Condition(4),"@SortType:"&Condition(5),"@RecorderCount:0","@pageSize:"&iPageSize,"@PageIndex:"&iPageIndex,"@@RecordCount","@@PageCount"))
						Case Else	'ʹ���Զ����ҳ�洢����
							spReturn = Array(False,False)
							For i = 0 To Ubound(Condition)
								If LCase(Condition(i)) = "@@recordcount" Then spReturn(0) = True
								If LCase(Condition(i)) = "@@pagecount" Then spReturn(1) = True
								If spReturn(0) And spReturn(1) Then Exit For
							Next
							If spReturn(0) And spReturn(1) Then
								spResult = doSP(pType&":3",Condition)
							Else
								ErrMsg "��ȡ��ҳ���ݳ���", "ʹ���Զ����ҳ�洢����ʱ�������@@RecordCount��@@PageCount���������"
							End If
					End Select
					Set GetPageRecord = spResult(0)
					iRecordCount = int(spResult(1)("@@RecordCount"))
					iPageCount = int(spResult(1)("@@PageCount"))
					iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
				Else
					ErrMsg "��ȡ��ҳ���ݳ���", "ʹ�ô洢���̻�ȡ��ҳ����ʱ������������Ϊ���飡"
				End If
		End Select
		If Instr(",sql,rs,1,2,", "," & pType & ",")>0 Then
			iRecordCount = rs.RecordCount
			rs.PageSize = iPageSize
			iPageCount = rs.PageCount
			iPageIndex = Easp_IIF(iPageIndex > iPageCount, iPageCount, iPageIndex)
			If iRecordCount>0 Then rs.AbsolutePage = iPageIndex
			Set GetPageRecord = rs
		End If
	End Function
	Public Function GPR(ByVal PageSetup, ByVal Condition)
		Set GPR = GetPageRecord(PageSetup, Condition)
	End Function
	'���ɷ�ҳ��������
	Public Function Pager(ByVal PagerHtml, ByRef PagerConfig)
		On Error Resume Next
		Dim pList, pListStart, pListEnd, pFirst, pPrev, pNext, pLast
		Dim pJump, pJumpLong, pJumpStart, pJumpEnd, pJumpValue
		Dim i, j, tmpStr, pStart, pEnd, cfg, pcfg(1)
		tmpStr = Easp_IIF(PagerHtml="",iPageDic("default_html"),PagerHtml)
		Set cfg = Server.CreateObject("Scripting.Dictionary")
		cfg("recordcount")	= iRecordCount
		cfg("pageindex")	= iPageIndex
		cfg("pagecount")	= iPageCount
		cfg("pagesize")		= iPageSize
		cfg("listlong")		= 9
		cfg("listsidelong")	= 2
		cfg("list")			= "*"
		cfg("currentclass")	= "current"
		cfg("link")			= GetRQ(0) & "*"
		cfg("first")		= "&laquo;"
		cfg("prev")			= "&#8249;"
		cfg("next")			= "&#8250;"
		cfg("last")			= "&raquo;"
		cfg("more")			= "..."
		cfg("disabledclass")= "disabled"
		cfg("jump")			= "input"
		cfg("jumpplus")		= ""
		cfg("jumpaction")	= ""
		cfg("jumplong")		= 50
		PagerConfig = Easp_IIF(isArray(PagerConfig),PagerConfig, Easp_IIF(Easp_isN(PagerConfig),iPageDic("default_config"),Array(PagerConfig,"pagerconfig:1")))
		If isArray(PagerConfig) Then
			Dim ConfigName, ConfigValue
			For i = 0 To Ubound(PagerConfig)
				ConfigName = LCase(Left(PagerConfig(i),Instr(PagerConfig(i),":")-1))
				ConfigValue = Mid(PagerConfig(i),Instr(PagerConfig(i),":")+1)
				If Instr(",recordcount,pageindex,pagecount,pagesize,listlong,listsidelong,jumplong,", ","&ConfigName&",") > 0 Then
					cfg(ConfigName) = Int(ConfigValue)
				Else
					cfg(ConfigName) = ConfigValue
				End If
			Next
		End If
		pStart = cfg("pageindex") - ((cfg("listlong") \ 2) + (cfg("listlong") Mod 2)) + 1
		pEnd = cfg("pageindex") + (cfg("listlong") \ 2)
		If pStart < 1 Then
			pStart = 1 : pEnd = cfg("listlong")
		End If
		If pEnd > cfg("pagecount") Then
			pStart = cfg("pagecount") - cfg("listlong") + 1 : pEnd = cfg("pagecount")
		End If
		If pStart < 1 Then pStart = 1
		For i = pStart To pEnd
			If i = cfg("pageindex") Then
				pList = pList & " <span class="""&cfg("currentclass")&""">" & Replace(cfg("list"),"*",i) & "</span> "
			Else
				pList = pList & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
			End If
		Next
		If cfg("listsidelong")>0 Then
			If cfg("listsidelong") < pStart Then
				For i = 1 To cfg("listsidelong")
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
				pListStart = pListStart & Easp_IIF(cfg("listsidelong")+1=pStart,"",cfg("more") & " ")
			ElseIf cfg("listsidelong") >= pStart And pStart > 1 Then
				For i = 1 To (pStart - 1)
					pListStart = pListStart & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
			If (cfg("pagecount") - cfg("listsidelong")) > pEnd Then
				pListEnd = " " & cfg("more") & pListEnd
				For i = ((cfg("pagecount") - cfg("listsidelong"))+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			ElseIf (cfg("pagecount") - cfg("listsidelong")) <= pEnd And pEnd < cfg("pagecount") Then
				For i = (pEnd+1) To cfg("pagecount")
					pListEnd = pListEnd & " <a href="""&Replace(cfg("link"),"*",i)&""">" & Replace(cfg("list"),"*",i) & "</a> "
				Next
			End If
		End If
		If cfg("pageindex") > 1 Then
			pFirst = " <a href="""&Replace(cfg("link"),"*","1")&""">" & cfg("first") & "</a> "
			pPrev = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")-1)&""">" & cfg("prev") & "</a> "
		Else
			pFirst = " <span class="""&cfg("disabledclass")&""">" & cfg("first") & "</span> "
			pPrev = " <span class="""&cfg("disabledclass")&""">" & cfg("prev") & "</span> "
		End If
		If cfg("pageindex") < cfg("pagecount") Then
			pLast = " <a href="""&Replace(cfg("link"),"*",cfg("pagecount"))&""">" & cfg("last") & "</a> "
			pNext = " <a href="""&Replace(cfg("link"),"*",cfg("pageindex")+1)&""">" & cfg("next") & "</a> "
		Else
			pLast = " <span class="""&cfg("disabledclass")&""">" & cfg("last") & "</span> "
			pNext = " <span class="""&cfg("disabledclass")&""">" & cfg("next") & "</span> "
		End If
		Select Case LCase(cfg("jump"))
			Case "input"
				pJumpValue = "this.value"
				pJump = "<input type=""text"" size=""3"" title=""������Ҫ��ת����ҳ�����س�""" & Easp_IIF(cfg("jumpplus")="",""," "&cfg("jumpplus"))
				pJump = pJump & " onkeydown=""javascript:if(event.charCode==13||event.keyCode==13){if(!isNaN(" & pJumpValue & ")){"
				pJump = pJump & Easp_IIF(cfg("jumpaction")="",Easp_IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*", pJumpValue))
				pJump = pJump & "}return false;}"" />"
			Case "select"
				pJumpValue = "this.options[this.selectedIndex].value"
				pJump = "<select" & Easp_IIF(cfg("jumpplus")="",""," "&cfg("jumpplus")) & " onchange=""javascript:"
				pJump = pJump & Easp_IIF(cfg("jumpaction")="",Easp_IIF(Lcase(Left(cfg("link"),11))="javascript:",Replace(Mid(cfg("link"),12),"*",pJumpValue),"document.location.href='" & Replace(cfg("link"),"*","'+" & pJumpValue & "+'") & "';"),Replace(cfg("jumpaction"),"*",pJumpValue))
				pJump = pJump & """ title=""��ѡ��Ҫ��ת����ҳ��""> "
				If cfg("jumplong")=0 Then
					For i = 1 To cfg("pagecount")
						pJump = pJump & "<option value=""" & i & """" & Easp_IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
					Next
				Else
					pJumpLong = Int(cfg("jumplong") / 2)
					pJumpStart = Easp_IIF(cfg("pageindex")-pJumpLong<1, 1, cfg("pageindex")-pJumpLong)
					pJumpStart = Easp_IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpStart-(pJumpLong-(cfg("pagecount")-cfg("pageindex")))+1, pJumpStart)
					pJumpStart = Easp_IIF(pJumpStart<1,1,pJumpStart)
					j = 1
					For i = pJumpStart To cfg("pageindex")
						pJump = pJump & "<option value=""" & i & """" & Easp_IIF(i=cfg("pageindex")," selected=""selected""","") & ">" & i & "</option> "
						j = j + 1
					Next
					pJumpLong = Easp_IIF(cfg("pagecount")-cfg("pageindex")<pJumpLong, pJumpLong, pJumpLong + (pJumpLong-j)+1)
					pJumpEnd = Easp_IIF(cfg("pageindex")+pJumpLong>cfg("pagecount"), cfg("pagecount"), cfg("pageindex")+pJumpLong)
					For i = cfg("pageindex")+1 To pJumpEnd
						pJump = pJump & "<option value=""" & i & """>" & i & "</option> "
					Next
				End If
				pJump = pJump & "</select>"
		End Select
		tmpStr = Replace(tmpStr,"{recordcount}",cfg("recordcount"))
		tmpStr = Replace(tmpStr,"{pagecount}",cfg("pagecount"))
		tmpStr = Replace(tmpStr,"{pageindex}",cfg("pageindex"))
		tmpStr = Replace(tmpStr,"{pagesize}",cfg("pagesize"))
		tmpStr = Replace(tmpStr,"{list}",pList)
		tmpStr = Replace(tmpStr,"{liststart}",pListStart)
		tmpStr = Replace(tmpStr,"{listend}",pListEnd)
		tmpStr = Replace(tmpStr,"{first}",pFirst)
		tmpStr = Replace(tmpStr,"{prev}",pPrev)
		tmpStr = Replace(tmpStr,"{next}",pNext)
		tmpStr = Replace(tmpStr,"{last}",pLast)
		tmpStr = Replace(tmpStr,"{jump}",pJump)
		Set cfg = Nothing
		Pager = vbCrLf & tmpStr & vbCrLf
	End Function
	'���÷�ҳ��ʽ
	Public Sub SetPager(ByVal PagerName, ByVal PagerHtml, ByRef PagerConfig)
		If PagerName = "" Then PagerName = "default"
		If Not Easp_isN(PagerHtml) Then iPageDic.item(PagerName&"_html") = PagerHtml
		If Not Easp_isN(PagerConfig) Then iPageDic.item(PagerName&"_config") = PagerConfig
	End Sub
	'���÷�ҳ��ʽ
	Public Function GetPager(ByVal PagerName)
		If PagerName = "" Then PagerName = "default"
		GetPager = Pager(iPageDic(PagerName&"_html"),iPageDic(PagerName&"_config"))
	End Function
	'ȡ�õ�ǰҳ��
	Private Function GetCurrentPage()
		Dim rqParam, thisPage : thisPage = 1
		rqParam = Request.QueryString(iPageParam)
		If isNumeric(rqParam) Then
			If Int(rqParam) > 0 Then thisPage = Int(rqParam)
		End If
		GetCurrentPage = thisPage
	End Function
	'���س�ȥҳ��ĵ�ǰURL����
	Private Function GetRQ(pageNumer)
		Dim tmpStr,rq : tmpStr = ""
		For Each rq In Request.QueryString()
			If rq<>iPageParam Then tmpStr = tmpStr & "&" & rq & "=" & Server.UrlEncode(Request.QueryString(rq))
		Next
		GetRQ = Request.ServerVariables("SCRIPT_NAME") & "?" & Easp_IIF(tmpStr="","",Mid(tmpStr,2)&"&") & iPageParam & "=" & Easp_IIF(pageNumer=0,"",pageNumer)
	End Function
End Class

'EasyASP������ͨ�ú�������
Private Function Easp_IIF(ByVal Cn, ByVal T, ByVal F)
	If Cn Then Easp_IIF = T Else Easp_IIF = F End If
End Function
Private Function Easp_Param(ByVal s)
	Dim arr(2),t : t = Instr(s,":")
	If t > 0 Then
		arr(0) = Left(s,t-1) : arr(1) = Mid(s,t+1)
	Else
		arr(0) = s : arr(1) = ""
	End If
	Easp_Param = arr
End Function
Private Function Easp_isN(ByVal str)
	Easp_isN = False
	Select Case VarType(str)
		Case vbEmpty, vbNull
			Easp_isN = True : Exit Function
		Case vbString
			If str="" Then Easp_isN = True : Exit Function
		Case vbObject
			If TypeName(str)="Nothing" Or TypeName(str)="Empty" Then Easp_isN = True : Exit Function
		Case vbArray,8194,8204,8209
			If Ubound(str)=-1 Then Easp_isN = True : Exit Function
	End Select
End Function
Private Function Easp_JsEncode(ByVal str)
	If Not isN(str) Then
		str = Replace(str,Chr(92),"\\")
		str = Replace(str,Chr(34),"\""")
		str = Replace(str,Chr(39),"\'")
		str = Replace(str,Chr(9),"\t")
		str = Replace(str,Chr(13),"\r")
		str = Replace(str,Chr(10),"\n")
		str = Replace(str,Chr(12),"\f")
		str = Replace(str,Chr(8),"\b")
	End If
	Easp_JsEncode = str
End Function
Private Function Easp_Escape(ByVal str)
	Dim i,c,a,s : s = ""
	If Easp_isN(str) Then Easp_Escape = "" : Exit Function
	For i = 1 To Len(str)
		c = Mid(str,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Easp_Escape = s
End Function
Private Function Easp_UnEscape(ByVal str)
	Dim x, s
	x = InStr(str,"%")
	s = ""
	Do While x>0
		s = s & Mid(str,1,x-1)
		If LCase(Mid(str,x+1,1))="u" Then
			s = s & ChrW(CLng("&H"&Mid(str,x+2,4)))
			str = Mid(str,x+6)
		Else
			s = s & Chr(CLng("&H"&Mid(str,x+1,2)))
			str = Mid(str,x+3)
		End If
		x=InStr(str,"%")
	Loop
	Easp_UnEscape = s & str
End Function
Private Function Easp_RandStr(ByVal length)
	Dim allowStr, i, p
	allowStr = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ"
	p = Easp_Param(length)
	If Not isN(p(1)) Then : allowStr = p(1) : length = p(0) : p = "" : End If
	For i = 1 To length
		Randomize() : Easp_RandStr = Easp_RandStr & Mid(allowStr, Int(Len(allowStr) * Rnd + 1), 1)
	Next
End Function
Private Function Easp_Rand(ByVal min, ByVal max)
    Randomize() : Easp_Rand = Int((max - min + 1) * Rnd + min)
End Function
Private Function Easp_Test(ByVal Str, ByVal Pattern)
	If Easp_IsN(Str) Then Easp_Test = False : Exit Function
	Dim Reg
	Set Reg = New RegExp
	Reg.IgnoreCase = True
	Reg.Global = True
	Reg.Pattern = Pattern
	Easp_Test = Reg.Test(CStr(Str))
	Set Reg = Nothing
End Function
Private Function Easp_Replace(ByVal Str, ByVal rule, Byval Result, ByVal isM)
	Dim tmpStr,Reg : tmpStr = Str
	If Not Easp_isN(Str) Then
		Set Reg = New Regexp
		Reg.Global = True
		Reg.IgnoreCase = True
		If isM = 1 Then Reg.Multiline = True
		Reg.Pattern = rule
		tmpStr = Reg.Replace(tmpStr,Result)
		Set Reg = Nothing
	End If
	Easp_Replace = tmpStr
End Function
'����ƥ��
Private Function Easp_Match(ByVal Str, ByVal rule)
	Dim Reg
	Set Reg = New Regexp
	Reg.Global = True
	Reg.IgnoreCase = True
	Reg.Pattern = rule
	Set Easp_Match = Reg.Execute(Str)
	Set Reg = Nothing
End Function
%>