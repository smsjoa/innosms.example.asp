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
	<a href="GetBalance.asp">�ܿ� �ݾ� ��ȸ</a><br>
	<a href="SendSMS.asp">SMS ����</a><br>
	<a href="SendLMS.asp">LMS ����</a><br>
	<a href="SendMMS.asp">MMS ����</a><br>
	<a href="GetMessage.asp">���� ���� ��ȸ</a><br>
	<a href="CancelReservation.asp">���� ���� ���</a><br><br>
 </body>
</html>
