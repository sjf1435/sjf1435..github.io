<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
' *** Logout the current user.
MM_Logout = CStr(Request.ServerVariables("URL")) & "?MM_Logoutnow=1"
If (CStr(Request("MM_Logoutnow")) = "1") Then
  Session.Contents.Remove("mm_name")
  Session.Contents.Remove("mm_level")
  MM_logoutRedirectPage = "login.asp"
  ' redirect with URL parameters (remove the "MM_Logoutnow" query param).
  if (MM_logoutRedirectPage = "") Then MM_logoutRedirectPage = CStr(Request.ServerVariables("URL"))
  If (InStr(1, UC_redirectPage, "?", vbTextCompare) = 0 And Request.QueryString <> "") Then
    MM_newQS = "?"
    For Each Item In Request.QueryString
      If (Item <> "MM_Logoutnow") Then
        If (Len(MM_newQS) > 1) Then MM_newQS = MM_newQS & "&"
        MM_newQS = MM_newQS & Item & "=" & Server.URLencode(Request.QueryString(Item))
      End If
    Next
    if (Len(MM_newQS) > 1) Then MM_logoutRedirectPage = MM_logoutRedirectPage & MM_newQS
  End If
  Response.Redirect(MM_logoutRedirectPage)
End If
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>�ޱ����ĵ�</title>
<link href="Config/style.css" rel="stylesheet" type="text/css">
</head>

<body>
<table width="100%" align="center">
  <tr> 
    <td height="30" align="center" class="bgcolor-left">��</td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right">
	<a href="<%= MM_Logout %>" target="_parent">��½</a></td>
  </tr>
  <% if cint(session("mm_level")) <> 1 then %>
   <% if cint(session("mm_level")) = 3 then %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="company.asp" target="main">��˾����</a></td>
  </tr>
  <% end if %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="bargain.asp" target="main">��ͬ����</a></td>
  </tr>
  <% else %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_add.asp" target="main">��ͬ¼��</a></td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_list.asp" target="main">��ͬ����</a></td>
  </tr>
  <% end if %>
  <% if cint(session("mm_level")) = 3 then %>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="b_type.asp" target="main">��ͬ���</a></td>
  </tr>
  <tr> 
    <td height="30" align="center" class="bgcolor-right"><a href="admin.asp" target="main">��Ա����</a></td>
  </tr>
  <% end if %>
</table>
<p>��</p>
<table width="100%" align="center">
  <tr> 
    <td height="30" align="center" class="bgcolor-left">������Ϣ</td>
  </tr>
  <tr> 
    <td height="100" align="center" class="bgcolor-right">
<h3><%=session("mm_name")%></h3>
      <p>��ӭ��Ĺ���<br>
        ��ļ�����: 
        <% 
	dim mm_level
	mm_level=session("mm_level")
	if mm_level = 1 then
	response.Write("¼��Ա")
	end if
	if mm_level = 2 then
	response.Write("���Ա")
	end if
	if mm_level = 3 then
	response.Write("��������Ա")
	end if
	%>
      </p></td>
  </tr>
</table>
<p>��</p>
</body>
</html>
