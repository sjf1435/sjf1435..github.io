<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level3.asp" -->
<%

if(request.querystring("a_id") <> "") then 
Command1__mmid = request.querystring("a_id")
else
response.Write("��������")
response.end 
end if

%>
<%

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_conn_bargain_STRING
Command1.CommandText = "DELETE FROM t_admin  WHERE a_id = " + Replace(Command1__mmid, "'", "''") + ""
Command1.CommandType = 1
Command1.CommandTimeout = 0
Command1.Prepared = true
Command1.Execute()

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<meta http-equiv="refresh" content="1;URL=admin.asp">
</head>

<body>
ɾ���ɹ����ȴ�����... 
</body>
</html>
