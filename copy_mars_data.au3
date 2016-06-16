#pragma compile(ProductVersion, 0.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Скрипт для копирования файлов Mars с суточными экг мониторами)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород - 31-555 - nn-admin@nnkk.budzdorov.su)
#pragma compile(ProductName, copy_mars_data)

AutoItSetOption("TrayAutoPause", 0)
AutoItSetOption("TrayIconDebug", 1)

#include <File.au3>
#include <FileConstants.au3>
#include <Date.au3>

#Region ==========================    Check for temp folder and create log    ==========================
Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $messageToSend = ""
Local $current_pc_name = @ComputerName
Local $tempFolder = StringSplit(@SystemDir, "\")[1] & "\Temp\"
Local $errStr = "===ERROR=== "
ConsoleWrite("Current_pc_name: " & $current_pc_name & @CRLF)

Local $logFilePath = @ScriptDir & "\copy_mars_data.log"
Local $logFile = FileOpen($logFilePath, $FO_OVERWRITE)
ToLog($current_pc_name)
ToLog(@CRLF & "---Check for temp folder and create log---")

If $logFile = -1 Then ToLog($errStr & "Cannot create log file at " & $tempFolder)
#EndRegion

#Region ==========================    Variables    ==========================
Local $iniFile = @ScriptDir & "\copy_mars_data.ini"
Local $generalSection = "general"
Local $mailSection = "mail"

Local $server_backup = ""
Local $login_backup = ""
Local $password_backup = ""
Local $to_backup = ""
Local $send_email_backup = "1"
#EndRegion

#Region ==========================    Reading the main settings    ==========================
ToLog(@CRLF & "---Reading the main settings---")

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

Local $source = IniRead($iniFile, $generalSection, "source", "")
Local $destination = IniRead($iniFile, $generalSection, "destination", "")
Local $destinationArchive = IniRead($iniFile, $generalSection, "destinationArchive", "")
Local $filesLimit = IniRead($iniFile, $generalSection, "filesLimit", "")
Local $moveIntoTheArchiveOlderThanDays = IniRead($iniFile, $generalSection, "moveIntoTheArchiveOlderThanDays", "")
Local $delaySeconds = IniRead($iniFile, $generalSection, "delaySeconds", "")
Local $mask = IniRead($iniFile, $generalSection, "mask", "")
#EndRegion

#Region ==========================    Check for the settings error    ==========================
ToLog(@CRLF & "---Check for the settings error---")
If $source = "" Then ToLog($errStr & "Cannot find key: source")
If $destination = "" Then ToLog($errStr & "Cannot find key: destination")
If $destinationArchive = "" Then ToLog($errStr & "Cannot find key: destinationArchive")
If $filesLimit = "" Then ToLog($errStr & "Cannot find key: filesLimit")
If $moveIntoTheArchiveOlderThanDays = "" Then ToLog($errStr & "Cannot find key: moveIntoTheArchiveOlderThanDays")
If $delaySeconds = "" Then ToLog($errStr & "Cannot find key: delaySeconds")
If $mask = "" Then ToLog($errStr & "Cannot find key: mask")

If StringInStr($messageToSend, $errStr) Then
   SendEmail()
   Exit
EndIf

ToLog("source: " & $source)
ToLog("destination: " & $destination)
ToLog("destinationArchive: " & $destinationArchive)
ToLog("filesLimit: " & $filesLimit)
ToLog("moveIntoTheArchiveOlderThanDays: " & $moveIntoTheArchiveOlderThanDays)
ToLog("delaySeconds: " & $delaySeconds)
ToLog("mask: " & $mask)
#EndRegion

#Region ==========================    MainLoop     ==========================
While True
   CheckData()
   If StringInStr($messageToSend, $errStr) Then SendEmail()
   $messageToSend = ""
   Sleep($delaySeconds * 1000)
WEnd
#EndRegion

#Region ==========================    Functions     ==========================
Func CheckData()
   ToLog(@CRLF & "---CheckingData---")
   ToLog("Source folder: " & $source)
   If Not FileExists($source) Then ToLog($errStr & "Source path doesn't exists: " & $source)
   If Not FileExists($destination) Then ToLog($errStr & "Destination path doesn't exists: " & $destination)
   If Not FileExists($destinationArchive) Then ToLog($errStr & "Destination archive path doesn't exists: " & $destinationArchive)
   If StringInStr($messageToSend, $errStr) Then Return

   Local $sourceFiles = _FileListToArray($source, $mask, $FLTA_FILES, True)
   If Not IsArray($sourceFiles) Then
	  ToLog("No files " & $mask & " found in folder: " & $sourceFiles)
	  Return
   EndIf

   Local $lastCheck = IniRead($iniFile, $generalSection, "lastCheck", "")
   If StringLen($lastCheck) <> 14 Or Not StringIsAlNum($lastCheck) Then
	  $lastCheck = @YEAR & @MON & @MDAY - 1 & @HOUR & @MIN & @SEC
	  ToLog("The last check has incorrect value and it will be set automatically")
   EndIf

   ToLog("The last check was at: " & StringLeft($lastCheck, 4) & "/" & StringMid($lastCheck, 5, 2) & _
	  "/" & StringMid($lastCheck, 7, 2) & " " & StringMid($lastCheck, 9, 2) & _
	  ":" & StringMid($lastCheck, 11, 2) & ":" & StringMid($lastCheck, 13, 2))

   Local $filesToCopy[0]
   For $i = 1 To $sourceFiles[0]
	  Local $fileTime = FileGetTime($sourceFiles[$i], $FT_CREATED)
	  Local $new = True
	  If Int(_ArrayToString($fileTime, "")) < Int($lastCheck) Then $new = False
	  If $new Then _ArrayAdd($filesToCopy, $sourceFiles[$i])
   Next

   If UBound($filesToCopy) Then
	  ToLog(UBound($filesToCopy) & " new file(s) was found")
   Else
	  ToLog("There is no new files")
	  Return
   EndIf

   Local $fileCounter = GetLastIndex($destination, $mask)

   ToLog("---Copying files to the destination folder---")
   ToLog("The destination folder: " & $destination)
   Local $copiedFileList[0]
   For $file in $filesToCopy
	  Local $newFileName = StringReplace($mask, "*", $fileCounter)
	  If Not FileCopy($file, $destination & $newFileName) Then
		 ToLog($errStr & "Cannot write: " & $destination & $newFileName)
	  Else
		 ToLog(StringReplace($file, $source, "..\") & " -> " & "..\" & $newFileName)
		 $fileCounter += 1
		 _ArrayAdd($copiedFileList, $newFileName)
	  EndIf
   Next

   If UBound($copiedFileList) Then
	  ToLog("Successfully copied " & UBound($copiedFileList) & " file(s)")
	  Local $tempFileName = _TempFile()
	  Local $message = _Now() & @CRLF & @CRLF & "В программу MARS добавлено новых исследований: " & UBound($copiedFileList) & _
		 @CRLF & @CRLF & "Список добавленных файлов: " & @CRLF & _ArrayToString($copiedFileList, @CRLF)
	  FileWrite($tempFileName, $message)
	  Local $notepad = Run("Notepad.exe " & $tempFileName)
	  If Not $notepad Then ToLog($errStr & "Cannot launch notepad.exe")
	  $lastCheck = @YEAR & @MON & @MDAY & @HOUR & @MIN & @SEC
	  If Not IniWrite($iniFile,$generalSection, "lastCheck", $lastCheck) Then ToLog($errStr & "Cannot update the lastCheck value in ini")
   Else
	  ToLog($errStr & "No one file has been copied")
	  Return
   EndIf

   If UBound($filesToCopy) > $filesLimit Then
	  ToLog($errStr & "New files quantity exceed limits")
	  ToLog($errStr & "Moving files to the archive will be skipped")
	  Return
   EndIf

   ToLog("---Checking limits in the destination folder---")
   Local $destinationFiles = _FileListToArray($destination, $mask, $FLTA_FILES, True)
   If IsArray($destinationFiles) Then
	  If $destinationFiles[0] > $filesLimit Then
		 Local $needToMove = $destinationFiles[0] - $filesLimit
		 ToLog("In the destination folder " & $needToMove & " file(s) above limit")
		 Local $now = @YEAR & "/" & @MON & "/" & @MDAY
		 Local $filesToArchive[0]

		 For $i = 1 To UBound($destinationFiles, $UBOUND_ROWS) - 1
			Local $fileTime = FileGetTime($destinationFiles[$i]);, $FT_CREATED)
			$fileTime = $fileTime[0] & "/" & $fileTime[1] & "/" & $fileTime[2]
			If _DateDiff("D", $fileTime, $now) >= $moveIntoTheArchiveOlderThanDays Then
			   _ArrayAdd($filesToArchive, $destinationFiles[$i])
			EndIf
		 Next

		 If Not UBound($filesToArchive) Then
			ToLog($errStr & "There is no older files to move to the archive")
			Return
		 EndIf

		 If UBound($filesToArchive) < $needToMove Then
			ToLog($errStr & "Files older than " & $moveIntoTheArchiveOlderThanDays & " day(s) less on " & _
			   $needToMove - UBound($filesToArchive) & " than need to be moved to the archive: " & $needToMove)
		 EndIf

		 ToLog("Moving old file(s) to the archive")
		 ToLog("The destination archive folder: " & $destinationArchive)
		 Local $movedFileList[0]
		 Local $archiveLastIndex = GetLastIndex($destinationArchive, $mask)
		 For $file in $filesToArchive
			Local $newFileName = StringReplace($mask, "*", $archiveLastIndex)
			If Not FileMove($file, $destinationArchive & "\" & $newFileName) Then
			   ToLog($errStr & "Cannot move file to the archive: " & $file)
			Else
			   ToLog(StringReplace($file, $destination, "..\") & " -> " & "..\" & $newFileName)
			   $archiveLastIndex += 1
			   _ArrayAdd($movedFileList, $newFileName)
			EndIf
		 Next

		 If Not UBound($movedFileList) Then
			ToLog($errStr & "No one file has been moved")
		 Else
			ToLog("Successfully moved " & UBound($movedFileList) & " file(s) to the archive")
		 EndIf
	  EndIf
   EndIf
EndFunc

Func GetLastIndex($path, $searchMask)
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
   _FileWriteLog($logFile, $message)
EndFunc

Func SendEmail()
   If Not $send_email Then
	  FileClose($logFile)
	  Return
   EndIf

   ToLog(@CRLF & "---Sending email---")
   If _INetSmtpMailCom($server, "Copy MARS data", $login, $to, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, "", "", "", $login, $password) <> 0 Then

	  _INetSmtpMailCom($server_backup, "Copy MARS data", $login_backup, $to_backup, _
		 $current_pc_name & ": error(s) occurred", _
		 $messageToSend, "", "", "", $login_backup, $password_backup)
   EndIf

   FileClose($logFile)
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
	  For $x = 1 To $S_Files2Attach[0] - 1
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