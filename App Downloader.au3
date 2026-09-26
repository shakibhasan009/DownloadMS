#RequireAdmin

; Register COM error handler to prevent hard script crashes on connection drops
Global $g_oCOMError = ObjEvent("AutoIt.Error", "_COMErrorTrap")

InstallMSStorePackage("9p95dll7rzz1")
InstallMSStorePackage("9msn9csm353x")
InstallMSStorePackage("9p91x732kd1z")
InstallMSStorePackage("9pp3g2bd3pdl")
InstallMSStorePackage("9phnb71mkr4j")
InstallMSStorePackage("9p95dll7rzz1")

;InstallMSStorePackage("9wzdncrfj39b")
;InstallMSStorePackage("9wzdncrcwpth")
;InstallMSStorePackage("9ntpjw96tsc0")
;InstallMSStorePackage("9mxgz0664zld")
;InstallMSStorePackage("9wzdncrfhwcp")
#cs
;DownloadMSStorePackage("9npqgc3lck31") ; Ink Draft[cite: 4]
;DownloadMSStorePackage("9p21845v8z53") ; Nina[cite: 4]
DownloadMSStorePackage("9p3xkpd2dg5g") ; Ink Journal[cite: 4]
DownloadMSStorePackage("9n9dzg1xt2mb") ; Georgia Pro[cite: 4]
DownloadMSStorePackage("9n8d67vhhdc2") ; Verdana Pro[cite: 4]
DownloadMSStorePackage("9pk93bg0z1jj") ; Gill Sans Nova[cite: 4]
DownloadMSStorePackage("9ns5ct1mz7m8") ; Arial Nova[cite: 4]
DownloadMSStorePackage("9n1wc5zsct3c") ; Convection[cite: 4]
DownloadMSStorePackage("9nd0mgdq2jdd") ; Monotype Ceremonial[cite: 4]
DownloadMSStorePackage("9p4vc9qlg8qj") ; Monotype Handwriting[cite: 4]
DownloadMSStorePackage("9ndsbn7gjxb0") ; Buxton Sketch[cite: 4]
DownloadMSStorePackage("9n57vdp26cd7") ; Rockwell Nova[cite: 4]
DownloadMSStorePackage("9nsbp8sgq3k3") ; Monotype Christmas[cite: 4]
DownloadMSStorePackage("9nxkc5g4bgvl") ; Clincher Mono[cite: 4]
DownloadMSStorePackage("9nfnb3hpjc84") ; Monotype Rockstar[cite: 4]
DownloadMSStorePackage("9msnbcq2jmkt") ; Monotype Halloween[cite: 4]
DownloadMSStorePackage("9ngqc45wnjb2") ; Segoe Marker[cite: 4]
DownloadMSStorePackage("9pg347fc4mp3") ; PT Root UI[cite: 4]
DownloadMSStorePackage("9nmfx2j3ckcw") ; Monotype Wild West[cite: 4]
DownloadMSStorePackage("9NPQCDPGJ6SZ")
DownloadMSStorePackage("9nblggh4vzw5")
DownloadMSStorePackage("9nblggh4wr1n")
DownloadMSStorePackage("9n8jxs6275rg")
DownloadMSStorePackage("9P5L31JC2M67")
#ce

; --- Function 1: Download Packages with Optional Target Directory ---
Func DownloadMSStorePackage($id, $targetDir = "")
	Local $pid = StringRegExp($id, '(?i)([a-z0-9]{12})', 1)
	If Not IsArray($pid) Then Return False

	Local $arch = (@OSArch = "X64") ? "x64" : "x86"
	Local $tmpHtml = @TempDir & "\meta_" & $pid[0] & ".html"
	Local $postData = "type=ProductId&url=" & $pid[0] & "&ring=RP&lang=en-US"

	Local $cmd = 'curl.exe -s -k -X POST "https://store.rg-adguard.net/api/GetFiles" ' & _
			'-H "Content-Type: application/x-www-form-urlencoded" ' & _
			'-H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64)" ' & _
			'-H "Referer: https://store.rg-adguard.net/" ' & _
			'--data "' & $postData & '" -o "' & $tmpHtml & '"'

	RunWait(@ComSpec & ' /c ' & $cmd, "", @SW_HIDE)
	If Not FileExists($tmpHtml) Then Return False

	Local $sResponse = FileRead($tmpHtml)
	FileDelete($tmpHtml)

	Local $m = StringRegExp($sResponse, '(?i)<a\s+[^>]*href="([^"]+)"[^>]*>([^<]+)</a>', 3)
	If Not IsArray($m) Then Return False

	Local $aDepU[1], $aDepF[1], $iDepCount = 0
	Local $appU = "", $appF = "", $appRank = -1

	For $i = 0 To UBound($m) - 1 Step 2
		Local $u = $m[$i], $f = $m[$i + 1]

		If StringInStr($f, ".BlockMap") Or StringInStr($f, "arm") Then ContinueLoop
		If StringRegExp($f, '(?i)\.resource-') And Not StringRegExp($f, '(?i)(neutral|_en)') Then ContinueLoop
		If Not (StringInStr($f, $arch) Or StringInStr($f, "neutral") Or StringRegExp($f, '(?i)\.(msixbundle|appxbundle)$')) Then ContinueLoop

		; Collect dependencies
		If StringRegExp($f, '(?i)(VCLibs|UI\.Xaml|NET\.Native).*\.(appx|msix)$') Then
			ReDim $aDepU[$iDepCount + 1]
			ReDim $aDepF[$iDepCount + 1]
			$aDepU[$iDepCount] = $u
			$aDepF[$iDepCount] = $f
			$iDepCount += 1
			ContinueLoop
		EndIf

		; Rank main package
		Local $r = StringRegExp($f, '(?i)\.msixbundle$') ? 4 : (StringRegExp($f, '(?i)\.msix$') ? 3 : (StringRegExp($f, '(?i)\.appxbundle$') ? 2 : (StringRegExp($f, '(?i)\.appx$') ? 1 : -1)))
		If $r > $appRank Then
			$appRank = $r
			$appU = $u
			$appF = $f
		EndIf
	Next

	If $appF = "" Then Return False

	; Determine destination directory based on optional parameter
	Local $sCleanName = CleanPackageName($appF)
	Local $finalDir = $targetDir
	If $finalDir = "" Then
		Local $baseDir = @UserProfileDir & "\Downloads\Programs\App Installer"
		$finalDir = $baseDir & "\" & $pid[0] & " - " & $sCleanName
	EndIf

	; Skip if target folder and all files exist
	Local $bAllExist = FileExists($finalDir) And FileExists($finalDir & "\" & $appF)
	If $bAllExist Then
		For $d = 0 To $iDepCount - 1
			If Not FileExists($finalDir & "\" & $aDepF[$d]) Then
				$bAllExist = False
				ExitLoop
			EndIf
		Next
	EndIf

	If $bAllExist Then
		ToolTip($sCleanName & " already downloaded. Skipping.", 5, 5, "Store Downloader", 1)
		Return True
	EndIf

	If Not FileExists($finalDir) Then DirCreate($finalDir)

	; Download dependencies
	For $d = 0 To $iDepCount - 1
		Local $depFile = $finalDir & "\" & $aDepF[$d]
		If Not FileExists($depFile) Then
			ToolTip('Downloading Runtime: ' & $aDepF[$d] & @CRLF & 'Please Wait...', 5, 5, "Store Downloader")
			InetGet($aDepU[$d], $depFile, 1, 0)
		EndIf
	Next

	; Download main app package
	Local $mainFile = $finalDir & "\" & $appF
	If Not FileExists($mainFile) Then
		ToolTip('Downloading: ' & $sCleanName & '...' & @CRLF & 'Please Wait...', 5, 5, "Store Downloader")
		InetGet($appU, $mainFile, 1, 0)
	EndIf

	ToolTip($sCleanName & " downloaded successfully!", 5, 5, "Store Downloader", 1)
	Return True
EndFunc   ;==>DownloadMSStorePackage

; --- Function 2: Direct Install Package with Skip If Installed ---
Func InstallMSStorePackage($id)
    Local $pid = StringRegExp($id, '(?i)([a-z0-9]{12})', 1)
    If Not IsArray($pid) Then Return False

    Local $dir = @TempDir & "\" & $pid[0], $arch = (@OSArch = "X64") ? "x64" : "x86"
    If Not FileExists($dir) Then DirCreate($dir)

    Local $postData = "type=ProductId&url=" & $pid[0] & "&ring=RP&lang=en-US"
    Local $tmpHtml = $dir & "\meta.html"
    Local $cmd = 'curl.exe -s -k -X POST "https://store.rg-adguard.net/api/GetFiles" ' & _
                 '-H "Content-Type: application/x-www-form-urlencoded" ' & _
                 '-H "User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64)" ' & _
                 '-H "Referer: https://store.rg-adguard.net/" ' & _
                 '--data "' & $postData & '" -o "' & $tmpHtml & '"'

    RunWait(@ComSpec & ' /c ' & $cmd, "", @SW_HIDE)
    If Not FileExists($tmpHtml) Then
        DirRemove($dir, 1)
        Return False
    EndIf

    Local $sResponse = FileRead($tmpHtml)
    FileDelete($tmpHtml)

    Local $m = StringRegExp($sResponse, '(?i)<a\s+[^>]*href="([^"]+)"[^>]*>([^<]+)</a>', 3)
    If Not IsArray($m) Then
        DirRemove($dir, 1)
        Return False
    EndIf

    Local $appU = "", $appF = "", $appRank = -1
    Local $depList = ""

    For $i = 0 To UBound($m) - 1 Step 2
        Local $u = $m[$i], $f = $m[$i + 1]

        If StringInStr($f, ".BlockMap") Or StringInStr($f, "arm") Then ContinueLoop
        If StringRegExp($f, '(?i)\.resource-') And Not StringRegExp($f, '(?i)(neutral|_en)') Then ContinueLoop
        If Not (StringInStr($f, $arch) Or StringInStr($f, "neutral") Or StringRegExp($f, '(?i)\.(msixbundle|appxbundle)$')) Then ContinueLoop

        ; Collect dependencies if not already installed
        If StringRegExp($f, '(?i)(VCLibs|UI\.Xaml|NET\.Native).*\.(appx|msix)$') Then
            If IsPackageInstalled($f) Then ContinueLoop

            Local $dPath = $dir & "\" & $f
            If Not FileExists($dPath) Then InetGet($u, $dPath, 1, 0)
            $depList &= ($depList = "" ? "" : ",") & "'" & $dPath & "'"
            ContinueLoop
        EndIf

        ; Rank main package
        Local $r = StringRegExp($f, '(?i)\.msixbundle$') ? 4 : (StringRegExp($f, '(?i)\.msix$') ? 3 : (StringRegExp($f, '(?i)\.appxbundle$') ? 2 : (StringRegExp($f, '(?i)\.appx$') ? 1 : -1)))
        If $r > $appRank Then
            $appRank = $r
            $appU = $u
            $appF = $f
        EndIf
    Next

    If $appF <> "" Then
        Local $cleanTitle = CleanPackageName($appF)

        If IsPackageInstalled($appF) Then
            ToolTip($cleanTitle & " is already installed. Skipping.", 5, 5, "Store Installer", 1)
            Sleep(1200)
            ToolTip("")
            DirRemove($dir, 1)
            Return True
        EndIf

        Local $mainPath = $dir & "\" & $appF

        ToolTip('Downloading ' & $cleanTitle & '...' & @CRLF & 'Please Wait...', 5, 5, "Store Downloader")
        If Not FileExists($mainPath) Then InetGet($appU, $mainPath, 1, 0)

        ToolTip('Installing ' & $cleanTitle & '...' & @CRLF & 'Please Wait...', 5, 5, "Store Installer")

        Local $dismDepArgs = ""
        If $depList <> "" Then
            Local $aDeps = StringSplit(StringReplace($depList, "'", ""), ",")
            For $j = 1 To $aDeps[0]
                $dismDepArgs &= ' /DependencyPackagePath:"' & $aDeps[$j] & '"'
            Next
        EndIf

        Local $exitCode = RunWait(@ComSpec & ' /c dism /Online /Add-ProvisionedAppxPackage /PackagePath:"' & $mainPath & '"' & $dismDepArgs & ' /SkipLicense', "", @SW_HIDE)

        If $exitCode <> 0 Then
            Local $psDep = ($depList <> "") ? " -DependencyPath @(" & $depList & ")" : ""
            RunWait(@ComSpec & ' /c powershell -NoP -ExecutionPolicy Bypass -Command "Add-AppxPackage -Path ''' & $mainPath & '''' & $psDep & '"', "", @SW_HIDE)
        EndIf

        DirRemove($dir, 1)
        ToolTip("")
        Return True
    EndIf

    DirRemove($dir, 1)
    Return False
EndFunc

; Intercept COM runtime errors so AutoIt won't crash
Func _COMErrorTrap($oError)
	Return
EndFunc   ;==>_COMErrorTrap

Func IsPackageInstalled($sFile)
	Local $a = StringRegExp($sFile, '^(?:[^._]+?\.)?([^._]+)_', 1)
	If Not IsArray($a) Then Return False
	Local $iExit = RunWait(@ComSpec & ' /c powershell -NoP -C "if(Get-AppxPackage -Name *' & $a[0] & '*){exit 0}else{exit 1}"', "", @SW_HIDE)
	Return ($iExit = 0)
EndFunc   ;==>IsPackageInstalled

Func CleanPackageName($sRaw)
	Local $a = StringRegExp($sRaw, '^(?:[^._]+?\.)?([^._]+)_([0-9.]+)', 1)
	If Not IsArray($a) Then Return $sRaw
	Local $name = StringRegExpReplace($a[0], '([a-z])([A-Z])', '$1 $2')
	$name = StringRegExpReplace($name, '([a-zA-Z])([0-9])', '$1 $2')
	$name = StringRegExpReplace($name, '([0-9])([a-zA-Z])', '$1 $2')
	Return StringStripWS($name & " " & $a[1], 3)
EndFunc   ;==>CleanPackageName
