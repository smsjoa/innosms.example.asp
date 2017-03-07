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
data.add "msg_serial", "¾¾¸®¾ó Å°"

Set ms = New MessageService
ms.getToken client_id, api_key

result = ms.cancelReservation(data)

Response.Write result

%>