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

data.add "phone", "수신번호_1" '한 명 전송
'data.add "phone", "수신번호_1, 수신번호_2" '여러 명 전송

'Set msg_list = jsObject() '개별 전송
'msg_list("수신번호_1") = "메시지_1" '개별 전송
'msg_list("수신번호_2") = "메시지_2" '개별 전송
'msg_list("수신번호_3") = "메시지_3" '개별 전송
'data.add "msg_list", toJSON(msg_list) '개별 전송

'data.add "trandate", "20150101000000" '예약 전송


Set ms = New MessageService
ms.getToken client_id, api_key

result = ms.sendMessage(data)

Set data = Nothing
Set msg_list = Nothing

Response.Write result

%>