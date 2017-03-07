<% @CODEPAGE="949" language="vbscript" %>
<% Option Explicit %>
<% session.CodePage = "949" %>
<% Response.CharSet = "euc-kr" %>
<% Response.buffer = True %>
<% Response.Expires = 0 %>

<!DOCTYPE html>
<html>
<head>
	<title>API Client</title>
</head>
 <body>
	<a href="GetBalance.asp">잔여 금액 조회</a><br>
	<a href="SendSMS.asp">SMS 전송</a><br>
	<a href="SendLMS.asp">LMS 전송</a><br>
	<a href="SendMMS.asp">MMS 전송</a><br>
	<a href="GetMessage.asp">전송 내역 조회</a><br>
	<a href="CancelReservation.asp">예약 내역 취소</a><br><br>
 </body>
</html>
