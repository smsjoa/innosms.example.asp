<% @CODEPAGE="65001" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim ms, result, data

Set data = Server.CreateObject("Scripting.Dictionary")
data.add "msg_serial", "������ Ű"
data.add "list_count", "������ ����"
data.add "page", "������ ��ȣ"

Set ms = New MessageService
ms.getToken client_id, api_key

result = ms.getMessage(data)

Response.Write result

%>