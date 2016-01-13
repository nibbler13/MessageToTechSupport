#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <ColorConstants.au3>
#include <Inet.au3>
#include <AD.au3>
#include <MsgBoxConstants.au3>

Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $user = @UserName
Local $userName = ""
Local $city = ""
Local $department = ""
Local $title = ""
Local $telephoneNumber = ""
Local $mail = ""
Local $physicalDeliveryOfficeName = ""
Local $userPrincipalName = ""
Local $arrPrinterList[1]

ReadUserInfo()

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("", 600, 450, -1, -1)
GUISetFont(10, $FW_NORMAL , 0, "Arial", $Form1, $CLEARTYPE_QUALITY)

$Label1 = GUICtrlCreateLabel("Обращение в службу технической поддержки клиники Будь Здоров", 10, 10, 580, 20, $SS_CENTER)
GUICtrlSetFont(-1, 11, $FW_SEMIBOLD)
GUICtrlSetColor(-1, "0x3030AA")
$Label5 = GUICtrlCreateLabel("внутренний номер: 603    для регионов: 30-494    городской: (495) 782-88-82", 10, 30, 580, 20, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_GRAY)

$Group1 = GUICtrlCreateGroup("Укажите информацию о себе:", 10, 70, 580, 124)
GUICtrlSetFont(-1, 10,  $FW_SEMIBOLD)

$Label2 = GUICtrlCreateLabel("Контактный номер телефона:", 20, 103, 180, 20, $SS_RIGHT)
$Input1 = GUICtrlCreateInput("", 210, 100, 363, 24)
If $telephoneNumber <> "" Then
   GUICtrlSetData(-1, $telephoneNumber)
EndIf

$Label3 = GUICtrlCreateLabel("Номер кабинета:", 20, 133, 180, 20, $SS_RIGHT)
$Input2 = GUICtrlCreateInput("", 210, 130, 363, 24)
If $physicalDeliveryOfficeName <> "" Then
   GUICtrlSetData(-1, $physicalDeliveryOfficeName)
EndIf

$Label4 = GUICtrlCreateLabel("ФИО:", 20, 163, 180, 20, $SS_RIGHT)
$Input3 = GUICtrlCreateInput("", 210, 160, 363, 24)
If $userName <> "" Then
   GUICtrlSetData(-1, $userName)
Else
   GUICtrlSetData(-1, $user)
EndIf

GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group2 = GUICtrlCreateGroup("Текст вашего обращения:", 10, 200, 580, 190)
GUICtrlSetFont(-1, 10,  $FW_SEMIBOLD)
$Edit1 = GUICtrlCreateEdit("", 20, 230, 560, 150, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
GUICtrlSetState(-1, $GUI_FOCUS)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$Button1 = GUICtrlCreateButton("Отправить обращение", 200, 400, 200, 40)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GetPrinterList()

While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		 Exit
	  Case $Button1
		 Local $tempString = ""

		 If GUICtrlRead ($Input1) = "" Then
			$tempString = $tempString & "Необходимо указать контактный номер телефона" & @CRLF
		 EndIf

		 If GUICtrlRead ($Input2) = "" Then
			$tempString = $tempString & "Необходимо указать номер кабинета" & @CRLF
		 EndIf


		 If GUICtrlRead ($Input3) = "" Then
			$tempString = $tempString & "Необходимо указать ФИО" & @CRLF
		 EndIf

		 If GUICtrlRead ($Edit1) = "" Then
			$tempString = $tempString & "Необходимо написать текст обращения" & @CRLF
		 EndIf

		 If $tempString <> "" Then
			MsgBox($MB_ICONWARNING, "", $tempString)
		 Else
			SendMessage()
		 EndIf
   EndSwitch
WEnd

Func SendMessage()
   Local $message = GUICtrlRead($Edit1) & @CRLF & @CRLF & @CRLF & "Инициатор:" & @TAB & @TAB & GUICtrlRead ($Input3) & @CRLF & _
					 "Контактный тел.:" & @TAB & GUICtrlRead ($Input1) & @CRLF & _
					 "Номер кабинета:" & @TAB & GUICtrlRead ($Input2) & @CRLF & @CRLF

   If $city <> "" Then
	  $message = $message & "Подразделение:" & @TAB & $city & @CRLF
   EndIf

   If $department <> "" Then
	  $message = $message & "Отдел:" & @TAB & @TAB & @TAB & $department & @CRLF
   EndIf

   If $title <> "" Then
	  $message = $message & "Должность:" & @TAB & @TAB & $title & @CRLF
   EndIf

   If $mail <> "" Then
	  $message = $message & "Почта:" & @TAB & @TAB & @TAB & $mail & @CRLF
   EndIf

   If $userPrincipalName <> "" Then
	  $message = $message & "Учетная запись:" & @TAB & $userPrincipalName & @CRLF
   EndIf

   $message = $message & @CRLF & "Имя компьютера:" & @TAB & @ComputerName & @CRLF & _
						"Имя пользователя:" & @TAB & @UserName & @CRLF & _
						"Версия ОС:" & @TAB & @TAB & @OSVersion & @CRLF & _
						"Архитектура ОС:" & @TAB & @OSArch & @CRLF & _
						"IP адрес:" & @TAB & @TAB & @IPAddress1 & @CRLF & _
						"Дата обращения:" & @TAB & @YEAR & "." & @MON & "." & @MDAY & "     " & @HOUR & ":" & @MIN

   If IsArray($arrPrinterList) Then
	  If UBound($arrPrinterList) > 1 Then
		 $message = $message & @CRLF & @CRLF & "Список установленных у пользователя принтеров:" & @CRLF
		 For $i = 1 To UBound($arrPrinterList) - 1
			$message = $message & $arrPrinterList[$i] & @CRLF
		 Next
	  EndIf
   EndIf


   Local $responce = _INetSmtpMailCom("172.16.6.6", "Auto request", "auto_request@7828882.ru", _
						"stp@7828882.ru", "Обращение через приложение STP", $message, "", "", "", _
						"auto_request@7828882.ru", "821973")

   Local $error = @error
   If $responce = 0 And $error = 0 Then
	  $tmp = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO), "", _
		 "Ваше обращение успешно отправлено!" & @CRLF & @CRLF & _
		 "Благодарим за обращение в службу технической поддержки." & @CRLF & _
		 "В ближайщее время с Вами свяжется специалист ИТ отдела," & @CRLF & _
		 "ответственный за выполнение данного обращения." & @CRLF & @CRLF & _
		 "Выйти из приложения?")
	  Switch $tmp
	  Case $IDYES
		 Exit
	  Case $IDNO
	  EndSwitch
   EndIf
EndFunc

Func ReadUserInfo()
   _AD_Open()

   If @error Then
	  Return
   EndIf

   $userName = _AD_GetObjectAttribute(@UserName, "displayName")
   $city = _AD_GetObjectAttribute(@UserName, "l")
   $department = _AD_GetObjectAttribute(@UserName, "department")
   $title = _AD_GetObjectAttribute(@UserName, "title")
   $telephoneNumber = _AD_GetObjectAttribute(@UserName, "telephoneNumber")
   $mail = _AD_GetObjectAttribute(@UserName, "mail")
   $physicalDeliveryOfficeName = _AD_GetObjectAttribute(@UserName, "physicalDeliveryOfficeName")
   $userPrincipalName = _AD_GetObjectAttribute(@UserName, "userPrincipalName")
   _AD_Close()
EndFunc

Func GetPrinterList()
   For $i = 1 To 1000
	  $reg = RegEnumVal( "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices", $i)
	  If @error = -1 Then ExitLoop
	  _ArrayAdd($arrPrinterList, $reg)
    Next

    $arrPrinterList[0] = $i - 1
EndFunc

;==============================================================================================================
; Description :  Send an email with a SMTP server by Microsoft CDO technology
; Parametere  : $s_SmtpServer
;               $s_FromName
;               $s_FromAddress
;               $s_ToAddress
;               $s_Subject
;               $as_Body
;               $s_AttachFiles (path file to join)
;               $s_CcAddress
;               $s_BccAddress
;               $s_Username
;               $s_Password
;               $IPPort
;               $ssl
; Return      :  On success none
;                On error code+msg
;==============================================================================================================
Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", $s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)

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

    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf

    ;Update settings
    $objEmail.Configuration.Fields.Update

    ; Sent the Message
    $objEmail.Send
    if @error then
        SetError(2)
        ;return $oMyRet[1]
    EndIf
EndFunc ;==>_INetSmtpMailCom
;===========================================================================================================

Func HandleComError()
   If $oMyError.source = "Active Directory" Then
	  ConsoleWrite("source is empty")
	  Return
   EndIf

   Msgbox($MB_ICONERROR, "Обращение в техподдержку", "Возникла ошибка при отправке." & @CRLF & _
			"Обратитесь в техподдержку по телефону: 603 или 30-494"    & @CRLF  & @CRLF & _
			"Описание ошибки: " & @TAB & $oMyError.description);  & @CRLF & _
			;"err.windescription:"   & @TAB & $oMyError.windescription & @CRLF & _
			;"err.number is: "       & @TAB & hex($oMyError.number,8)  & @CRLF & _
			;"err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
			;"err.scriptline is: "   & @TAB & $oMyError.scriptline   & @CRLF & _
			;"err.source is: "       & @TAB & $oMyError.source       & @CRLF & _
			;"err.helpfile is: "       & @TAB & $oMyError.helpfile     & @CRLF & _
			;"err.helpcontext is: " & @TAB & $oMyError.helpcontext _
            ;)
Endfunc