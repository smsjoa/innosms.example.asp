<%

Function Base64Encode(ByVal asContents)
	Dim sBASE_64_CHARACTERS, lnPosition, lsResult, Char1, Char2, Char3, Char4, Byte1, Byte2, Byte3, SaveBits1, SaveBits2, lsGroupBinary, lsGroup64

	sBASE_64_CHARACTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

	If Len(asContents) Mod 3 > 0 Then asContents = asContents & String(3 - (Len(asContents) Mod 3), " ")

	lsResult = ""

	For lnPosition = 1 To Len(asContents) Step 3
		lsGroup64 = ""
		lsGroupBinary = Mid(asContents, lnPosition, 3)

		Byte1 = Asc(Mid(lsGroupBinary, 1, 1)): SaveBits1 = Byte1 And 3
		Byte2 = Asc(Mid(lsGroupBinary, 2, 1)): SaveBits2 = Byte2 And 15
		Byte3 = Asc(Mid(lsGroupBinary, 3, 1))

		Char1 = Mid(sBASE_64_CHARACTERS, ((Byte1 And 252) \ 4) + 1, 1)
		Char2 = Mid(sBASE_64_CHARACTERS, (((Byte2 And 240) \ 16) Or (SaveBits1 * 16) And &HFF) + 1, 1)
		Char3 = Mid(sBASE_64_CHARACTERS, (((Byte3 And 192) \ 64) Or (SaveBits2 * 4) And &HFF) + 1, 1)
		Char4 = Mid(sBASE_64_CHARACTERS, (Byte3 And 63) + 1, 1)

		lsGroup64 = Char1 & Char2 & Char3 & Char4

		lsResult = lsResult + lsGroup64
	Next

	Base64Encode = lsResult
End Function


Function createBinaryRequestBody(requestBody, imageData, requestEnd)
	Const adTypeBinary = 1
	Const adModeReadWrite = 3

	Dim adoStream, headData, tailData, messageData

	headData = strToBinary(requestBody)
	tailData = strToBinary(requestEnd)

	Set adoStream = Server.CreateObject("ADODB.Stream")

	adoStream.Type = adTypeBinary
	adoStream.Mode = adModeReadWrite
	adoStream.Open
	adoStream.Write headData
	adoStream.Write imageData
	adoStream.Write tailData
	adoStream.Position = 0

	messageData = adoStream.Read

	adoStream.Close

	Set adoStream = Nothing

	createBinaryRequestBody = messageData
End Function


Function strToBinary(toConvert)
	Const adTypeBinary = 1
	Const adTypeText = 2
	Const adModeReadWrite = 3

	Dim adoStream, data

	Set adoStream = Server.CreateObject("ADODB.Stream")

	adoStream.Charset = "utf-8"
	adoStream.Type = adTypeText
	adoStream.Mode = adModeReadWrite

	adoStream.Open
	adoStream.WriteText toConvert

	adoStream.Position = 0
	adoStream.Type = adTypeBinary
	adoStream.Position = 3
	data = adoStream.Read

	adoStream.Close

	Set adoStream = Nothing

	strToBinary = data
End Function


Function getBoundary()
	Const boundaryLength = 20
	Const usableChars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789"

	Dim boundary, randIndex, i

	Randomize

	boundary = ""

	For i=1 To boundaryLength
		randIndex = Int(Len(usableChars) * Rnd + 1)
		boundary = boundary & Right(Left(usableChars, randIndex), 1)
	Next

	getBoundary = "------------------------" & boundary

End Function


Function createHTTPTextRequestParam(boundary, key, value)
	Dim message

	message = boundary & vbCrLf & "Content-Disposition: form-data; name=""" & key & """" & vbCrLf & vbCrLf & value & vbCrLf
	createHTTPTextRequestParam = message
End Function


Function readImage(imageLocation)
	Const adTypeBinary = 1
	Dim adoStream, imageData

	Set adoStream = Server.CreateObject("ADODB.Stream")

	adoStream.Type = adTypeBinary
	adoStream.Open
	adoStream.LoadFromFile imageLocation
	imageData = adoStream.Read
	adoStream.Close

	Set adoStream = Nothing

	readImage = imageData
End Function

%>