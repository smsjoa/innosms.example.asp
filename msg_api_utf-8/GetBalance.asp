<% @CODEPAGE="65001" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "65001" %>
<% Response.CharSet = "utf-8" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!-- #include file="lib/MessageService.asp" -->
<!-- #include file="common.asp" -->

<%

Dim ms

Set ms = New MessageService
ms.getToken client_id, api_key

Response.Write ms.getBalance()

%>