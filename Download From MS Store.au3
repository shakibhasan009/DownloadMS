#cs ----------------------------------------------------------------------------

 Shakib Hasan
 Email: ShakibHasan726@Gmail.com


#ce ----------------------------------------------------------------------------

Func DownloadMS($sID)
	Global $apiUrl = "https://store.rg-adguard.net/api/GetFiles"
	Global $productUrl = "https://apps.microsoft.com/detail/" & $sID
	Global $downloadFolder = @TempDir & "\" & $sID

	If Not FileExists($downloadFolder) Then DirCreate($downloadFolder)

	Global $body = "type=url&url=" & $productUrl & "&ring=RP&lang=en-US"
	Global $oInet = ObjCreate("winhttp.winhttprequest.5.1")
	$oInet.Open("POST", $apiUrl, False)
	$oInet.SetRequestHeader("Content-Type", "application/x-www-form-urlencoded")
	$oInet.Send($body)
	Global $response = $oInet.ResponseText


	Global $matches = StringRegExp($response, '<tr style.*<a href="(?<url>.*?)"\s.*>(?<text>.*?)</a>', 3)
	If IsArray($matches) Then
		For $i = 0 To UBound($matches)
			Local $url = $matches[$i]
			Local $name = $matches[Int($i) + 1]
			ToolTip($name, 5, 5, "Downloading From Microsoft Store")
			If @OSArch = "x86" Then
				If Not StringInStr($name, ".BlockMap") Then
					If Not StringInStr($name, "_arm") Then
						If Not StringInStr($name, "x64") Then InetGet($url, $downloadFolder & "\" & $name)
					EndIf
				EndIf
			ElseIf @OSArch = "x64" Then
				If Not StringInStr($name, ".BlockMap") Then
					If Not StringInStr($name, "_arm") Then InetGet($url, $downloadFolder & "\" & $name)
				EndIf
			Else
				If Not StringInStr($name, ".BlockMap") Then InetGet($url, $downloadFolder & "\" & $name)
			EndIf
			$i += 1
		Next
	Else
		DirRemove($downloadFolder)
	EndIf
EndFunc

DownloadMS("9WZDNCRFHVJL") ; OneNote for Windows 10
DownloadMS("9nn6g50qnl1h") ; Angry Birds Friends