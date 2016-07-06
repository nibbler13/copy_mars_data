#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#pragma compile(ProductVersion, 0.71)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для копирования файлов Mars с суточными экг мониторами)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555)
#pragma compile(ProductName, copy_mars_data)

AutoItSetOption("TrayAutoPause", 0)
AutoItSetOption("TrayIconDebug", 1)

#include <File.au3>
#include <FileConstants.au3>
#include <Date.au3>
#include <Crypt.au3>
#include <String.au3>

#Region ==========================    Variables    ==========================
Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $messageToSend = ""
Local $current_pc_name = @ComputerName
Local $errStr = "===ERROR=== "
ConsoleWrite("Current_pc_name: " & $current_pc_name & @CRLF)

Local $logFilePath = @ScriptDir & "\" & @ScriptName & "_" & @YEAR & @MON & @MDAY & ".log"
If FileExists($logFilePath) Then FileDelete($logFilePath)
Local $historyFileName = "Журнал добавления исследований MARS.txt"
Local $historyFilePath = @ScriptDir & "\" & $historyFileName
ToLog($current_pc_name)

Local $iniFile = @ScriptDir & "\copy_mars_data.ini"
Local $generalSection = "general"
Local $sourcesSection = "sources"
Local $mailSection = "mail"

Local $server_backup = ""
Local $login_backup = ""
Local $password_backup = ""
Local $to_backup = ""
Local $send_email_backup = "1"

Local $server = IniRead($iniFile, $mailSection, "server", $server_backup)
Local $login = IniRead($iniFile, $mailSection, "login", $login_backup)
Local $password = IniRead($iniFile, $mailSection, "password", $password_backup)
Local $to = IniRead($iniFile, $mailSection, "to", $to_backup)
Local $send_email = IniRead($iniFile, $mailSection, "send_email", $send_email_backup)

If $send_email = "" Then $send_email = "1"

If Not FileExists($iniFile) Then
   ToLog($errStr & "Cannot find settings file: " & $iniFile)
   SendEmail()
EndIf

ToLog("server: " & $server)
ToLog("login: " & $login)
ToLog("to: " & $to)
ToLog("send_mail: " & $send_email)

Local $source = IniReadSection($iniFile, $sourcesSection)
Local $destination = IniRead($iniFile, $generalSection, "destination", "")
Local $delaySeconds = IniRead($iniFile, $generalSection, "delaySeconds", "")
Local $mask = IniRead($iniFile, $generalSection, "mask", "")

Local $destinationHashKeys
#EndRegion

#Region ==========================    Check for the settings error    ==========================
ToLog(@CRLF & "---Check for settings errors---")
If Not IsArray($source) Then ToLog($errStr & "Cannot find section: sources")
If $destination = "" Then ToLog($errStr & "Cannot find key: destination")
If $delaySeconds = "" Then ToLog($errStr & "Cannot find key: delaySeconds")
If $mask = "" Then ToLog($errStr & "Cannot find key: mask")

If StringInStr($messageToSend, $errStr) Then
   SendEmail()
   Exit
EndIf

ToLog("source: " & _ArrayToString($source, " "))
ToLog("destination: " & $destination)
ToLog("delaySeconds: " & $delaySeconds)
ToLog("mask: " & $mask)
#EndRegion

#Region ==========================    MainLoop     ==========================
While True
   If UBound($source, $UBOUND_ROWS) Then
		Local $messageToUser = ""

		Local $lastCheck = IniRead($iniFile, $generalSection, "lastCheck", "")
		If StringLen($lastCheck) <> 14 Or Not StringIsAlNum($lastCheck) Then
			$lastCheck = @YEAR & @MON & @MDAY - 1 & @HOUR & @MIN & @SEC
			ToLog("The last check has incorrect value and it will be set automatically")
		EndIf

		ToLog("The last check was at: " & StringLeft($lastCheck, 4) & "/" & StringMid($lastCheck, 5, 2) & _
			"/" & StringMid($lastCheck, 7, 2) & " " & StringMid($lastCheck, 9, 2) & _
			":" & StringMid($lastCheck, 11, 2) & ":" & StringMid($lastCheck, 13, 2))

		_Crypt_Startup()
		$destinationHashKeys = GetHashKey($destination, $mask)

		For $i = 1 To UBound($source, $UBOUND_ROWS) - 1
			$messageToUser &= CheckData($source[$i][1], $lastCheck, $source[$i][0])
		Next
		_Crypt_Shutdown()

		If $messageToUser <> "" Then
			$messageToUser = _Now() & @CRLF & "ВНИМАНИЕ! Добавлены новые исследования суточного мониторирования" & @CRLF & _
				"Необходимо перезапустить программу (MARS)" & @CRLF & $messageToUser
			Local $tempFileName = _TempFile()
			FileWrite($tempFileName, $messageToUser)
			Local $notepad = Run("Notepad.exe " & $tempFileName)
			If Not $notepad Then ToLog($errStr & "Cannot launch notepad.exe")
			Local $historyFileLink = @DesktopDir & "\" & $historyFileName & ".lnk"
			If Not FileExists($historyFileLink) Then _
				FileCreateShortcut($historyFilePath, $historyFileLink)
		EndIf
	Else
		ToLog($errStr & "The source key doesn't contain any path")
	EndIf

	If StringInStr($messageToSend, $errStr) Then
	   SendEmail()
	Else
		Local $lastCheck = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC
		If Not IniWrite($iniFile,$generalSection, "lastCheck", $lastCheck) Then _
			ToLog($errStr & "Cannot update the lastCheck value in ini")
	EndIf

   $messageToSend = ""
   ToLog(@CRLF & "---Sleeping " & $delaySeconds & " second(s)")
   Sleep($delaySeconds * 1000)
WEnd
#EndRegion

#Region ==========================    Functions     ==========================
Func CheckData($path, $lastCheck, $displayName)
	ToLog(@CRLF & "---CheckingData---")
	ToLog("Source folder: " & $path)

	Local $message = ""
	If Not FileExists($path) Then ToLog($errStr & "Source path doesn't exists: " & $path)
	If Not FileExists($destination) Then ToLog($errStr & "Destination path doesn't exists: " & $destination)
	If StringInStr($messageToSend, $errStr) Then Return

	Local $sourceFiles = _FileListToArray($path, $mask, $FLTA_FILES, True)
	If Not IsArray($sourceFiles) Then
		ToLog("Didn't found files " & $mask & " in folder")
		Return $message
	EndIf

	Local $filesToCopy[0]
	Local $copiedFileList[0][2]
	Local $fileCounter = GetLastIndex($destination, $mask)

	For $i = 1 To $sourceFiles[0]
		Local $currentFilePath = $sourceFiles[$i]
		ToLog("Current file: " & $currentFilePath)

		Local $fileTime = FileGetTime($currentFilePath, $FT_MODIFIED, $FT_STRING)
		ToLog(@TAB & "Modified time: " & $fileTime)
		If Int($fileTime) < Int($lastCheck) Then
			ToLog(@TAB & "Skipping - too old")
			ContinueLoop
		EndIf

		Local $fileHash = _Crypt_HashFile($currentFilePath, $CALG_MD5)
		ToLog(@TAB & "File hash: " & $fileHash)
		If _ArraySearch($destinationHashKeys, $fileHash) > -1 Then
			ToLog(@TAB & "Skipping - hash key already present")
			ContinueLoop
		EndIf

		_ArrayAdd($destinationHashKeys, $fileHash)
		Local $newFileName = StringReplace($mask, "*", $fileCounter)

		If Not FileCopy($currentFilePath, $destination & $newFileName) Then
			ToLog($errStr & "Cannot write: " & $destination & $newFileName)
		Else
			ToLog(@TAB & "Copying the file: " & $currentFilePath)
			$fileCounter += 1

			Local $file = FileOpen($destination & $newFileName, BitOR($FO_ANSI, $FO_READ))
;~ 			Local $file = FileOpen($currentFilePath, BitOR($FO_ANSI, $FO_READ))
			Local $fullName = "Имя неизвестно"
			Local $stringWithName = ""
			Local $line = 1

			While True
				$stringWithName = FileReadLine($file, $line)
				If @error = 1 Or @error = -1 Then ExitLoop
				If StringInStr($stringWithName, "PtRace") Then ExitLoop
				If $line > 500 Then ExitLoop
				$line += 1
			WEnd

			FileClose($file)

			If StringInStr($stringWithName, "PtRace") And _
				StringInStr($stringWithName, "PtLName") And _
				StringInStr($stringWithName, "PtGender") And _
				StringInStr($stringWithName, "PtFName") Then

				Local $result[0]

				Local $ascii = StringToASCIIArray($stringWithName)
				For $symbol = 0 To UBound($ascii) - 1
					If $ascii[$symbol] Then _ArrayAdd($result, $ascii[$symbol])
				Next
				$stringWithName = StringFromASCIIArray($result)
;~ 				ToLog($stringWithName)

				Local $n1start = StringInStr($stringWithName, "PtRace", $STR_CASESENSE) + 6
				Local $n1count = StringInStr($stringWithName, "PtLName", $STR_CASESENSE) - $n1start
				Local $n2start = StringInStr($stringWithName, "PtGender", $STR_CASESENSE) + 8
				Local $n2count = StringInStr($stringWithName, "PtFName", $STR_CASESENSE) - $n2start

				Local $patientName = StringMid($stringWithName, $n1start, $n1count) & " " & StringMid($stringWithName, $n2start, $n2count)
				$patientName = StringReplace($patientName, " ", "  ")
				$patientName = DeleteEvenSymbols($patientName)

				If $patientName Then $fullName = $patientName

;~ 				Local $n3Start = StringInStr($stringWithName, "RefMdFName", $STR_CASESENSE) + 10
;~ 				Local $n3count = StringInStr($stringWithName, "RecSerNum", $STR_CASESENSE) - $n3Start
;~ 				Local $n4Start = StringInStr($stringWithName, "PtLName", $STR_CASESENSE) + 7
;~ 				Local $n4count = StringInStr($stringWithName, "PtId", $STR_CASESENSE) - $n4Start

;~ 				Local $recSerNum = StringMid($stringWithName, $n3Start, $n3count)
;~ 				Local $ptId = StringMid($stringWithName, $n4Start, $n4count)

;~ 				ToLog("=== " & $recSerNum & " " & $ptId & " ===")
			EndIf

			ToLog(@TAB & StringReplace($currentFilePath, $path, "..\") & " -> " & "..\" & $newFileName & " | " & $fullName)

			Local $toAdd[1][2]
			$toAdd[0][0] = $fullName
			$toAdd[0][1] = $newFileName
			_ArrayAdd($copiedFileList, $toAdd)

			If Not _FileWriteLog($historyFilePath, $displayName & " | " & $newFileName & " | " & _
				$fullName & " | " & $fileTime & " | " & $fileHash, 1) Then _
				ToLog($errStr & "Cannot write to the history file: " & $historyFilePath)
		EndIf

		_ArrayAdd($filesToCopy, $currentFilePath)
	Next

	If Not UBound($filesToCopy, $UBOUND_ROWS) Then
		ToLog("There is no new files")
		Return $message
	EndIf

	ToLog(UBound($filesToCopy) & " new file(s) was found")
	ToLog("The destination folder: " & $destination)

	If Not UBound($copiedFileList, $UBOUND_ROWS) Then
		ToLog($errStr & "No files has been copied")
		Return $message
	EndIf

	_ArraySort($copiedFileList)
	$copiedFileList = NormalizeNameLength($copiedFileList)
	ToLog("Successfully copied " & UBound($copiedFileList) & " file(s)")

	Local $length = StringLen($copiedFileList[0][0] & " | " & $copiedFileList[0][1])
	Local $line = _StringRepeat("-", $length)
	$displayName &= _StringRepeat(" ", StringLen($copiedFileList[0][0]) - StringLen($displayName))
	Local $message =  @CRLF & $line & @CRLF & $displayName & " | " & _
		UBound($copiedFileList) & " шт." & @CRLF & $line & @CRLF & _
		_ArrayToString($copiedFileList, " | ", Default, Default, @CRLF) & @CRLF & $line & @CRLF

	Return $message
EndFunc

Func NormalizeNameLength($copiedFileList)
	Local $size = UBound($copiedFileList, $UBOUND_ROWS)
	If $size < 2 Then Return $copiedFileList

	Local $maxLength = 0
	For $i = 0 To $size - 1
		Local $length = StringLen($copiedFileList[$i][0])
		If $length > $maxLength Then $maxLength = $length
	Next

	For $i = 0 To $size - 1
		Local $length = StringLen($copiedFileList[$i][0])
		If $length < $maxLength Then $copiedFileList[$i][0] &= _StringRepeat(" ", $maxLength - $length)
	Next

	Return $copiedFileList
EndFunc

Func DeleteEvenSymbols($str)
	Local $tmp = ""
	For $i = 1 To StringLen($str)
		$tmp &= StringMid($str, $i, 1)
		$i += 1
	Next
	Return $tmp
EndFunc

Func GetHashKey($path, $searchMask)
	ToLog("---Calculating hash keys for files in: " & $path & "---")
	Local $result[0]
	Local $files = _FileListToArray($path, $searchMask, $FLTA_FILES, True)
	If Not IsArray($files) Or UBound($files) = 0 Then Return $result

	For $i = 1 To UBound($files) - 1
		Local $currentHash = _Crypt_HashFile($files[$i], $CALG_MD5)
		_ArrayAdd($result, $currentHash)
		ToLog($files[$i] & " | " & $currentHash)
	Next

	Return $result
EndFunc

Func GetLastIndex($path, $searchMask)
	ToLog("---Searching last index in: " & $path & "---")
   Local $files = _FileListToArray($path, $searchMask, $FLTA_FILES, True)
   If Not IsArray($files) Or UBound($files) = 0 Then Return 0

   _ArrayColInsert($files, 1)

   For $i = 1 To UBound($files, $UBOUND_ROWS) - 1
	  Local $fileName = $files[$i][0]
	  Local $maskLeft = StringLeft($searchMask, StringInStr($searchMask, "*") - 1)
	  Local $maskRight = StringRight($searchMask, StringLen($searchMask) - StringInStr($searchMask, "*"))
	  Local $start = StringInStr($fileName, $maskLeft) + StringLen($maskLeft)
	  Local $count = StringInStr($fileName, $maskRight)
	  $files[$i][1] = Int(StringMid($fileName, $start, $count - $start))
   Next

   Return _ArrayMax($files, 0, 1, Default, 1) + 1
EndFunc

Func ToLog($message)
   $message &= @CRLF
   $messageToSend &= $message
   ConsoleWrite($message)
   _FileWriteLog($logFilePath, $message)
EndFunc

Func SendEmail()
   If Not $send_email Then Return

   ToLog(@CRLF & "---Sending email---")
   If _INetSmtpMailCom($server, "Copy MARS data", $login, $to, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, $logFilePath, "", "", $login, $password) <> 0 Then

	  _INetSmtpMailCom($server_backup, "Copy MARS data", $login_backup, $to_backup, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, $logFilePath, "", "", $login_backup, $password_backup)
   EndIf
EndFunc

Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, _
   $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", _
   $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

   Local $objEmail = ObjCreate("CDO.Message")
   Local $i_Error = 0
   Local $i_Error_desciption = ""

   $objEmail.From = '"' & $s_FromName & '" <' & $s_FromAddress & '>'
   $objEmail.To = $s_ToAddress

   If $s_CcAddress <> "" Then $objEmail.Cc = $s_CcAddress
   If $s_BccAddress <> "" Then $objEmail.Bcc = $s_BccAddress

   $objEmail.Subject = $s_Subject

   If StringInStr($as_Body,"<") and StringInStr($as_Body,">") Then
	  $objEmail.HTMLBody = $as_Body
   Else
	  $objEmail.Textbody = $as_Body & @CRLF
   EndIf

   If $s_AttachFiles <> "" Then
	  Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
	  For $x = 1 To $S_Files2Attach[0]
		 $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
		 If FileExists($S_Files2Attach[$x]) Then
			$objEmail.AddAttachment ($S_Files2Attach[$x])
		 Else
			$i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
			SetError(1)
			return 0
		 EndIf
	  Next
   EndIf

   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
   $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort

   If $s_Username <> "" Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
   EndIf

   If $Ssl Then
	  $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
   EndIf

   $objEmail.Configuration.Fields.Update
   $objEmail.Send

   if @error then
	  SetError(2)
   EndIf

   Return @error
EndFunc

Func HandleComError()
   ToLog($errStr & @ScriptName & " (" & $oMyError.scriptline & ") : ==> COM Error intercepted!" & @CRLF & _
            @TAB & "err.number is: " & @TAB & @TAB & "0x" & Hex($oMyError.number) & @CRLF & _
            @TAB & "err.windescription:" & @TAB & $oMyError.windescription & @CRLF & _
            @TAB & "err.description is: " & @TAB & $oMyError.description & @CRLF & _
            @TAB & "err.source is: " & @TAB & @TAB & $oMyError.source & @CRLF & _
            @TAB & "err.helpfile is: " & @TAB & $oMyError.helpfile & @CRLF & _
            @TAB & "err.helpcontext is: " & @TAB & $oMyError.helpcontext & @CRLF & _
            @TAB & "err.lastdllerror is: " & @TAB & $oMyError.lastdllerror & @CRLF & _
            @TAB & "err.scriptline is: " & @TAB & $oMyError.scriptline & @CRLF & _
            @TAB & "err.retcode is: " & @TAB & "0x" & Hex($oMyError.retcode) & @CRLF & @CRLF)
Endfunc
#EndRegion