<% @CODEPAGE="949" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "949" %>
<% Response.CharSet = "euc-kr" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim ms, result, data, msg_list

Set data = Server.CreateObject("Scripting.Dictionary")
data.add "msg_type", "sms"
data.add "callback", ""
data.add "msg", ""

data.add "phone", "���Ź�ȣ_1" '�� �� ����
'data.add "phone", "���Ź�ȣ_1, ���Ź�ȣ_2" '���� �� ����

'Set msg_list = jsObject() '���� ����
'msg_list("���Ź�ȣ_1") = "�޽���_1" '���� ����
'msg_list("���Ź�ȣ_2") = "�޽���_2" '���� ����
'msg_list("���Ź�ȣ_3") = "�޽���_3" '���� ����
'data.add "msg_list", toJSON(msg_list) '���� ����

'data.add "trandate", "20150101000000" '���� ����


Set ms = New MessageService
ms.getToken client_id, api_key

result = ms.sendMessage(data)

Set data = Nothing
Set msg_list = Nothing

Response.Write result

%>