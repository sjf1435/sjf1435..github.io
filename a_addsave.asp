<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="Connections/conn_bargain.asp" -->
<!--#include file="chk_login.asp" -->
<!--#include file="chk_level3.asp" -->
<%
Dim rsa__MMColParam
rsa__MMColParam = "1"
If (Request.Form("a_username") <> "") Then 
  rsa__MMColParam = Request.Form("a_username")
End If
%>
<%
Dim rsa
Dim rsa_numRows

Set rsa = Server.CreateObject("ADODB.Recordset")
rsa.ActiveConnection = MM_conn_bargain_STRING
rsa.Source = "SELECT * FROM t_admin WHERE a_username = '" + Replace(rsa__MMColParam, "'", "''") + "'"
rsa.CursorType = 0
rsa.CursorLocation = 2
rsa.LockType = 1
rsa.Open()

rsa_numRows = 0
%>
<% 

If Not rsa.EOF Or Not rsa.BOF Then 
response.Write("�û����Ѿ�����")
response.end
End If ' end Not rsa.EOF Or NOT rsa.BOF 

%>
<%

if(request.form("a_level") <> "") then Command1__mmlevel = request.form("a_level")

if(request.form("a_password") <> "") then Command1__mmpassword = request.form("a_password")

if(request.form("a_username") <> "") then Command1__mmusername = request.form("a_username")

%>
<%

set Command1 = Server.CreateObject("ADODB.Command")
Command1.ActiveConnection = MM_conn_bargain_STRING
Command1.CommandText = "INSERT INTO t_admin (a_level, a_password, a_username)  VALUES ('" + Replace(Command1__mmlevel, "'", "''") + "','" + Replace(Command1__mmpassword, "'", "''") + "','" + Replace(Command1__mmusername, "'", "''") + "' ) "
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
����û��ɹ����ȴ�����... 
</body>
</html>
<%
rsa.Close()
Set rsa = Nothing
%>
