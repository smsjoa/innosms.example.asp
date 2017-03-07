
<script runat="server" language="javascript" src="JSON2.min.asp"></script>

<script language="javascript" runat="server">
	function decodeUTF8(str){
		return decodeURIComponent(str);
	}

	function encodeUTF8(str){
		return encodeURIComponent(str);
	}
</script>


<!-- #include file="JSON_2.0.4.asp" -->
<!-- #include file="function.asp" -->

<%

Class MessageService

	Private serviceUrl
	Private version
	Private token


	Private Sub Class_Initialize
		serviceUrl = "https://api.smsjoa.com"
		version = "1"
	End Sub

	Private Sub Class_Terminate
		deleteToken()
	End Sub


	Private Function execute(uri, method, header, contents, isPost, isMultiPart)
		Dim http, formData, headerKeys, headerItems, i, result

		Set http = Server.CreateObject("MSXML2.ServerXMLHTTP")

		If isMultiPart Then
			Dim boundary, image, param, requestBody, requestEnd, imageData
			boundary = getBoundary
			image = contents.item("image")
			contents.Remove("image")

			requestBody = ""

			For Each param in contents.Keys
				requestBody = requestBody & createHTTPTextRequestParam("--" & boundary, param, contents(param))
			Next

			requestBody = requestBody & "--" & boundary & vbCrLf & "Content-Disposition: form-data; name=""image""; filename=""" & image & """" & vbCrLf & "Content-Type: application/octet-stream" & vbCrLf & vbCrLf

			requestEnd = vbCrLf & "--" & boundary & "--" & vbCrLf
			imageData = readImage(image)
			formData = createBinaryRequestBody(requestBody, imageData, requestEnd)

			header.add "Content-Type", "multipart/form-data; boundary=" & boundary
		Else
			If Not isNull(contents) Then
				Dim contentKeys, contentItems, j, temp

				contentKeys = contents.Keys
				contentItems = contents.Items

				temp = ""

				For j = 0 To contents.Count - 1
					If isPost Then
						temp = temp & contentKeys(j) & "=" & encodeUTF8(contentItems(j)) & "&"
					Else
						temp = temp & "/" & contentItems(j)
					End If
				Next

				If isPost Then
					formData = temp
				Else
					uri = uri & temp
				End If

			End If

			header.add "Content-Type", "application/x-www-form-urlencoded; charset=euc-kr"
		End If

		http.open method, serviceUrl & "/" & version & "/" & uri, FALSE

		headerKeys = header.Keys
		headerItems = header.Items

		For i = 0 To header.Count - 1
			http.setRequestHeader headerKeys(i), headerItems(i)
		Next

		http.send formData

		result = http.responseText

		Set http = Nothing
		Set header = Nothing

		execute = result
	End Function


	Public Sub getToken(clientId, apiKey)
		Dim header, result

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Basic " & Base64Encode(clientId & ":" & apiKey)
		result = execute("token", "POST", header, NULL, FALSE, FALSE)

		token = JSON.parse(result).token
	End Sub


	Public Sub deleteToken()
		Dim header, result

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Bearer " & token
		result = execute("token", "DELETE", header, NULL, FALSE, FALSE)
	End Sub


	Public Function getBalance()
		Dim header, result

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Bearer " & token
		result = execute("balance", "GET", header, NULL, TRUE, FALSE)

		getBalance = JSON.parse(result).money
	End Function


	Public Function sendMessage(contents)
		Dim header, result, isMultiPart

		If contents.item("msg_type") = "mms" Then
			isMultiPart = TRUE
		Else
			isMultiPart = FALSE
		End If

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Bearer " & token
		result = execute("send", "POST", header, contents, TRUE, isMultiPart)

		sendMessage = result
	End Function


	Public Function getMessage(contents)
		Dim header, result

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Bearer " & token
		result = execute("send", "GET", header, contents, FALSE, FALSE)

		getMessage = result
	End Function


	Public Function cancelReservation(contents)
		Dim header, result

		Set header = Server.CreateObject("Scripting.Dictionary")
		header.add "Authorization", "Bearer " & token
		result = execute("reservation", "DELETE", header, contents, FALSE, FALSE)

		cancelReservation = result
	End Function

End Class


%>