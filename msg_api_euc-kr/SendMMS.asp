<% @CODEPAGE="949" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "949" %>
<% Response.CharSet = "euc-kr" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim ms, result, data

Set data = Server.CreateObject("Scripting.Dictionary")
data.add "msg_type", "mms"
data.add "callback", ""
data.add "subject", ""
data.add "msg", ""
data.add "image", "C:\mms\sample.jpg"

'data.add "trandate", "20150101000000" '���� ����

data.add "phone", "���Ź�ȣ_1" '�� �� ����
'data.add "phone", "���Ź�ȣ_1, ���Ź�ȣ_2" '���� �� ����

Set ms = New MessageService
ms.getToken client_id, api_key

result = ms.sendMessage(data)

Set data = Nothing

Response.Write result

%>