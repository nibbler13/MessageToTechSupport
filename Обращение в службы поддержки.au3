#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#pragma compile(ProductVersion, 3.0)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Программа для создания и отправки обращения в службы поддержки)
#pragma compile(LegalCopyright, Грашкин Павел Павлович - Нижний Новгород)
#pragma compile(ProductName, Обращение в службы поддержки)
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ***


#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <FontConstants.au3>
#include <ColorConstants.au3>
#include <Inet.au3>
#include <AD.au3>
#include <File.au3>
#include <MsgBoxConstants.au3>
#include <GuiListView.au3>
#include <ListBoxConstants.au3>
#include <WindowsConstants.au3>
#include <WinAPIEx.au3>
#Include <ScreenCapture.au3>
#Include <Misc.au3>
#include <GUIScrollbars_Ex.au3>
#include <GuiImageList.au3>
#include <GuiButton.au3>
#include <StringSize.au3>
#include <ComboConstants.au3>


Opt("MustDeclareVars", 1)


#Region ============== Variables =============
Local $file1 = "picComputerSupport.ico"
Local $file2 = "picBuildingSupport.ico"
Local $file3 = "picMedicalSupport.ico"

FileInstall("picComputerSupport.ico", @TempDir & "\" & $file1, $FC_OVERWRITE)
FileInstall("picBuildingSupport.ico", @TempDir & "\" & $file2, $FC_OVERWRITE)
FileInstall("picMedicalSupport.ico", @TempDir & "\" & $file3, $FC_OVERWRITE)

Local $color1 = 0x4265ad
Local $color2 = 0x314142
Local $color3 = 0xe75d4a

Local $textColor = 0x293c42

Local $oMyError = ObjEvent("AutoIt.Error","HandleComError")
Local $user = @UserName
Local $userName = ""
Local $city = ""
Local $department = ""
Local $title = ""
Local $phoneNumber = ""
Local $mail = ""
Local $physicalDeliveryOfficeName = ""
Local $userPrincipalName = ""
Local $arrPrinterList[1]
Local $listView = ""
Local $attachmentsListGroup = ""
Local $deleteButton = -666
Local $__aGUIDropFiles = 0
Local $iX1 = 0
Local $iY1 = 0
Local $iX2 = 0
Local $iY2 = 0
Local $__MonitorList[1][5]
Local $inX = 0
Local $inY = 0
Local $monWid = 0
Local $monHei = 0
Local $screenShotList[0]
Local $totalSize = 0
Local $emailTo = ""
Local $winHeader = ""
Local $icon = ""
Local $color = ""
Local $exitingMessage = ""

Local $headerTitleColor = $textColor

Local $dX = 600
Local $dY = 494

Local $gapSize = 10

Local $fontMainSize = 11
Local $fontTitleSize = 13
Local $fontName = "Arial"

Local $currentX = $gapSize * 3
Local $currentY = $gapSize * 3

Local $buttonWidth = Round(($dX - $gapSize * 10) / 4)
Local $buttonHeight = 100

Local $mainGui = 0
Local $phoneNumberInput = 0
Local $cabinetInput = 0
Local $fullNameInput = 0
Local $ticketTextEdit = 0
Local $criticalCheckbox = 0
Local $filesAddButton = 0
Local $screenshotButton = 0
Local $sendButton = 0

Local $iconWidth = 77
Local $iconHeight = $iconWidth
#EndRegion


FormDepartmentSelectGui()


Func FormDepartmentSelectGui()
	$mainGui = GUICreate("", $dX, $dY)

	GUISetFont($fontMainSize, $FW_NORMAL, -1, $fontName, $mainGui, $CLEARTYPE_QUALITY)

	Local $selectTitleLabel = GUICtrlCreateLabel("Выберите отдел, в который адресовано Ваше обращение:", _
		$currentX, $currentY, $dX - $currentX * 2, -1, BitOR($SS_CENTERIMAGE, $SS_CENTER))
	GUICtrlSetFont(-1, $fontTitleSize)
	GUICtrlSetColor(-1, $headerTitleColor)

	Local $str1 = "          Техническая поддержка информационных систем        "
	Local $str2 = "    Эксплуатация и тех. обслуживание зданий и сооружений   "
	Local $str3 = "  Сервисное обсулуживание медицинского оборудования "

	UpdateCurrentY($selectTitleLabel)
	$currentY -= $gapSize

	Local $startY = $currentY + ($dY - $currentY - 100 * 3 - $gapSize * 4) / 2

	Local $selectBut1 = GUICtrlCreateButton($str1, $currentX, $startY, $dX - $currentX * 2, $buttonHeight, $BS_MULTILINE)
	_GUICtrlButton_SetImage(-1, @TempDir & "\" & $file1)

	UpdateCurrentY($selectBut1)
	$currentY += $gapSize * 2

	Local $selectBut2 = GUICtrlCreateButton($str2, $currentX, $currentY, $dX - $currentX * 2, $buttonHeight, $BS_MULTILINE)
	_GUICtrlButton_SetImage(-1, @TempDir & "\" & $file2)

	UpdateCurrentY($selectBut2)
	$currentY += $gapSize * 2

	Local $selectBut3 = GUICtrlCreateButton($str3, $currentX, $currentY, $dX - $currentX * 2, $buttonHeight, $BS_MULTILINE)
	_GUICtrlButton_SetImage(-1, @TempDir & "\" & $file3)


	GUISetState(@SW_SHOW)

	ReadUserInfo()

	While 1
		Local $nMsg = GUIGetMsg()
		Switch $nMsg
		Case $selectBut1
			$winHeader = "Обращение в отдел технической поддержки" & @CRLF & "информационных систем"
			$icon = $file1
			$color = $color1
			$exitingMessage = "Благодарим за обращение в службу технической поддержки." & @CRLF & _
								"В ближайщее время с Вами свяжется специалист ИТ отдела," & @CRLF & _
								"ответственный за выполнение данного обращения." & @CRLF & @CRLF
			ExitLoop

		Case $selectBut2
			$winHeader = "Обращение в отдел эксплуатации и технического" & @CRLF & "обслуживания зданий и сооружений"
			$icon = $file2
			$color = $color2
			$exitingMessage = "Благодарим за обращение в службу эксплуатации здания." & @CRLF & _
								"В ближайщее время с Вами свяжется инженер по эксплуатации," & @CRLF & _
								"ответственный за выполнение данного обращения." & @CRLF & @CRLF
			ExitLoop

		Case $selectBut3
			$winHeader = "Обращение в отдел сервисного обслуживания" & @CRLF & "медицинского оборудования"
			$icon = $file3
			$color = $color3
			$emailTo = "sos@7828882.ru"
			$exitingMessage = "Благодарим за обращение в службу обслуживания медицинского оборудования." & @CRLF & _
								"В ближайщее время с Вами свяжется инженер по медицинскому оборудованию," & @CRLF & _
								"ответственный за выполнение данного обращения." & @CRLF & @CRLF
			ExitLoop

		Case $GUI_EVENT_CLOSE
			Exit
		EndSwitch
	WEnd

	GUISetState(@SW_LOCK)
	GUICtrlDelete($selectTitleLabel)
	GUICtrlDelete($selectBut1)
	GUICtrlDelete($selectBut2)
	GUICtrlDelete($selectBut3)

	GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")
	GUIRegisterMsg($WM_DROPFILES, 'WM_DROPFILES')

	FormMainGui()
EndFunc


Func FormMainGui()
	;------- Drag'n'Drop -------;
	Local $mainGroup = GUICtrlCreateGroup("", -15, -15, $dX + 30, $dY + 30)
	GUICtrlSetState(-1, $GUI_DROPACCEPTED)

	$currentX = $gapSize
	$currentY = $gapSize

	;------- Header -------;
	Local $mainIcon = GUICtrlCreateIcon(@TempDir & "\" & $icon, -1, $currentX, $currentY, $iconWidth, $iconHeight)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	$currentX += $iconWidth + $gapSize
	Local $mainTitleLabel = CreateLabel($winHeader, $currentX, $currentY, _
		$dX - $currentX - $gapSize, $iconHeight, $color, 0xFFFFFF, $fontTitleSize)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	UpdateCurrentY($mainIcon)
	$currentX = $gapSize * 2

	;------- Phone number -------;
	Local $phoneNumberLabel = GUICtrlCreateLabel("Контактный номер телефона:", $currentX, $currentY, -1, -1, $SS_RIGHT)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	GUICtrlSetColor(-1, $textColor)
	UpdateCurrentX($phoneNumberLabel)
	$phoneNumberInput = GUICtrlCreateInput("", $currentX, $currentY - 3, $dX - $currentX - $gapSize * 2, -1)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	If $phoneNumber <> "" Then GUICtrlSetData(-1, $phoneNumber)
	UpdateCurrentY($phoneNumberLabel)
	Local $prevCtrlPos = ControlGetPos($mainGui, "", $phoneNumberLabel)

	;------- Cabinet number -------;
	Local $cabinetLabel = GUICtrlCreateLabel("Номер кабинета:", $prevCtrlPos[0], $currentY, $prevCtrlPos[2], $prevCtrlPos[3], $SS_RIGHT)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	GUICtrlSetColor(-1, $textColor)
	$prevCtrlPos = ControlGetPos($mainGui, "", $phoneNumberInput)
	$cabinetInput = GUICtrlCreateInput("", $prevCtrlPos[0], $currentY - 3, $prevCtrlPos[2], $prevCtrlPos[3])
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	If $physicalDeliveryOfficeName <> "" Then GUICtrlSetData(-1, $physicalDeliveryOfficeName)
	UpdateCurrentY($cabinetLabel)
	$prevCtrlPos = ControlGetPos($mainGui, "", $cabinetLabel)

	;------- Full name -------;
	Local $fullNameLabel = GUICtrlCreateLabel("ФИО:", $prevCtrlPos[0], $currentY, $prevCtrlPos[2], $prevCtrlPos[3], $SS_RIGHT)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	GUICtrlSetColor(-1, $textColor)
	$prevCtrlPos = ControlGetPos($mainGui, "", $cabinetInput)
	$fullNameInput = GUICtrlCreateInput("", $prevCtrlPos[0], $currentY - 3, $prevCtrlPos[2], $prevCtrlPos[3])
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	If $userName <> "" Then
	   GUICtrlSetData(-1, $userName)
	Else
	   GUICtrlSetData(-1, $user)
	EndIf
	UpdateCurrentY($fullNameLabel)
	$currentX = $gapSize * 2

	;------- Ticket text -------;
	Local $ticketTextLabel = GUICtrlCreateLabel("Текст вашего обращения:", $currentX, $currentY, $dX - $gapSize * 4, -1)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	GUICtrlSetColor(-1, $textColor)
	UpdateCurrentY($ticketTextLabel)
	$currentY -= $gapSize
	$ticketTextEdit = GUICtrlCreateEdit("", $currentX, $currentY, $dX - $gapSize * 4, 160, _
		BitOR($ES_AUTOVSCROLL, $ES_WANTRETURN, $WS_VSCROLL), $WS_EX_CLIENTEDGE)
	GUICtrlSetResizing(-1, $GUI_DOCKALL)
	UpdateCurrentY($ticketTextEdit)

	;------- Critical checkbox -------;
	$criticalCheckbox = GUICtrlCreateCheckbox("Критичный приоритет (остановлена работа отделения)", $currentX, $currentY)
	$prevCtrlPos = ControlGetPos($mainGui, "", $criticalCheckbox)
	ControlMove($mainGui, "", $criticalCheckbox, $currentX + ($dX - $currentX - $gapSize * 2 - $prevCtrlPos[2]) / 2, _
		$currentY - ($prevCtrlPos[3] - $gapSize * 2) / 2)
	GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKWIDTH, $GUI_DOCKHEIGHT))
	GUICtrlSetColor(-1, $textColor)
	UpdateCurrentY($criticalCheckbox)
	$currentX = $gapSize * 2

	;------- Files add button -------;
	$filesAddButton = GUICtrlCreateButton("Вложить файл", $currentX, $currentY, $buttonWidth, 50, $BS_MULTILINE)
	GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKSIZE))
;~ 	GUICtrlSetFont(-1, $fontMainSize)

	;------- Screenshot button -------;
	$prevCtrlPos = ControlGetPos($mainGui, "", $filesAddButton)
	$screenshotButton = GUICtrlCreateButton("Сделать снимок экрана", $prevCtrlPos[0] + $prevCtrlPos[2] + $gapSize * 2, _
		$prevCtrlPos[1], $buttonWidth, $prevCtrlPos[3], $BS_MULTILINE)
	GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKSIZE))
;~ 	GUICtrlSetFont(-1, $fontMainSize)

	;------- Send button -------;
	$sendButton = GUICtrlCreateButton("Отправить обращение", _
		$dX - $buttonWidth - $gapSize * 2, $currentY, $buttonWidth, 50, $BS_MULTILINE)
	GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKSIZE))
	GUICtrlSetFont(-1, $fontMainSize, $FW_SEMIBOLD)

	GetPrinterList()

	GUICtrlSetState($ticketTextEdit, $GUI_FOCUS)
	GUISetStyle(-1, $WS_EX_ACCEPTFILES)

	GUICtrlCreateGroup("", -99, -99, 1, 1)
	GUISetState(@SW_UNLOCK)


	While 1
		Local $nMsg = GUIGetMsg()
		Switch $nMsg
			Case $GUI_EVENT_CLOSE
				DeleteScreenshots()
				Exit
			Case $sendButton
				SendButtonPressed()
			Case $screenshotButton
				ScreenshotButtonPressed()
			Case $filesAddButton
				FilesAddButtonPressed()
			Case $deleteButton
				DeleteButtonPressed()
			Case $GUI_EVENT_DROPPED
				GuiEventDropped()
			Case $criticalCheckbox
				CriticalCheckboxPressed()
		EndSwitch
		Sleep(20)
	WEnd
EndFunc




Func CreateLabel($text, $x, $y, $width, $height, $textColor, $backgroundColor, $fontSize)
	GUISetFont($fontSize)
	Local $label = GUICtrlCreateLabel($text, 0, 0, -1, -1, $SS_CENTER)
	GUICtrlSetBkColor(-1, $GUI_BKCOLOR_TRANSPARENT)
	GUICtrlSetColor(-1, $textColor)

	Local $position = ControlGetPos($mainGui, "", $label)
	If IsArray($position) Then
		Local $newX = $x + ($width - $position[2]) / 2
		Local $newY = $y + ($height - $position[3]) / 2
		GUICtrlSetPos($label, $newX, $newY)
	EndIf
	GUISetFont($fontMainSize)
EndFunc


Func UpdateCurrentX($controlID)
	Local $previousControlPosition = ControlGetPos($mainGui, "", $controlID)
	$currentX = $previousControlPosition[0] + $previousControlPosition[2] + $gapSize
EndFunc


Func UpdateCurrentY($controlID)
	Local $previousControlPosition = ControlGetPos($mainGui, "", $controlID)
	$currentY = $previousControlPosition[1] + $previousControlPosition[3] + $gapSize
EndFunc






Func ScreenshotButtonPressed()
   GUISetState(@SW_HIDE, $mainGui)

   $__MonitorList[0][0] = 0
   Local $handle = DllCallbackRegister("_MonitorEnumProc", "int", "hwnd;hwnd;ptr;lparam")
   DllCall("user32.dll", "int", "EnumDisplayMonitors", "hwnd", 0, "ptr", 0, "ptr", _
	  DllCallbackGetPtr($handle), "lparam", 0)
   DllCallbackFree($handle)

   Local $i = 0
   $__MonitorList[0][1] = 0
   $__MonitorList[0][2] = 0
   $__MonitorList[0][3] = 0
   $__MonitorList[0][4] = 0
   For $i = 1 To $__MonitorList[0][0]
	  If $__MonitorList[$i][1] < $__MonitorList[0][1] Then $__MonitorList[0][1] = $__MonitorList[$i][1]
	  If $__MonitorList[$i][2] < $__MonitorList[0][2] Then $__MonitorList[0][2] = $__MonitorList[$i][2]
	  If $__MonitorList[$i][3] > $__MonitorList[0][3] Then $__MonitorList[0][3] = $__MonitorList[$i][3]
	  If $__MonitorList[$i][4] > $__MonitorList[0][4] Then $__MonitorList[0][4] = $__MonitorList[$i][4]
   Next

   $inX = $__MonitorList[0][1]
   $inY = $__MonitorList[0][2]
   $monWid = $__MonitorList[0][3] + Abs($__MonitorList[0][1])
   $monHei = $__MonitorList[0][4]

   Local $Form3 = GUICreate("", $monWid, $monHei, $inX, $inY, $WS_POPUP, $WS_EX_TOPMOST)
   GUISetBkColor(0xFFFFFF)
   WinSetTrans($Form3, "", 100)
   GUISetState(@SW_SHOW, $Form3)
   GUISetCursor(3, 1, $Form3)

   Local $Form2 = GUICreate("Создание снимка экрана", 400, 56, -1, -1, -1, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST, $mainGui)
   GUISetFont($fontMainSize, $FW_NORMAL, -1, $fontName, $Form2, $CLEARTYPE_QUALITY)

   Local $cancelButton = GUICtrlCreateButton("Отмена", 8, 8, 78, 40)
   GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKLEFT, $GUI_DOCKSIZE))

   Local $contButton = GUICtrlCreateButton("Продолжить", 278, 8, 114, 40)
   GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKRIGHT, $GUI_DOCKSIZE))
   GUICtrlSetFont(-1, Default, $FW_HEAVY)
   GUICtrlSetState(-1, $GUI_HIDE)
	GUICtrlSetState($contButton, $GUI_FOCUS)

   Local $labelMouse = GUICtrlCreateLabel("Выделите мышью нужную область экрана", _
	  86, 19, 314, 40, $SS_CENTER)
   GUICtrlSetFont(-1, Default, $FW_HEAVY)

   GUISetState(@SW_SHOW, $Form2)

   Local $exit = False
   Local $screenCaptured = False
   Local $shotPath

   While Not $exit
	  Local $nMsg2 = GUIGetMsg()

	  If $nMsg2 ==  $GUI_EVENT_PRIMARYDOWN And Not $screenCaptured Then
		 Local $mPos = MouseGetPos()
		 Local $mX = $mPos[0]
		 Local $mY = $mPos[1]
		 Local $wPos = WinGetPos($Form2)
		 Local $wX = $wPos[0]
		 Local $wY = $wPos[1]
		 Local $wW = $wPos[2]
		 Local $wH = $wPos[3]

		 If Not ($mX >= $wX And $mX <= $wX + $wW And $mY >= $wY And $mY <= $wY + $wH) Then
			GUISetState(@SW_HIDE, $Form2)
			$iX1 = $mX
			$iY1 = $mY

			Mark_Rect()

			If Abs($iX1 - $iX2) < 10 Or Abs($iY1 - $iY2) < 10 Then
			   GUICtrlSetData($labelMouse, "Выделена слишком маленькая область экрана, попробуйте снова")
			   GUICtrlSetPos($labelMouse, 86, 10, 314, 40)
			   GUISetState(@SW_SHOW, $Form2)
			   ContinueLoop
			EndIf

			$shotPath = @TempDir & "\Снимок_экрана_" & @YEAR & "-" & @MON & "-" & @MDAY & _
			   "_" & @HOUR & "-" & @MIN & "-" & @SEC & ".jpg"

			GUIDelete($Form3)

			_ScreenCapture_Capture($shotPath, $iX1, $iY1, $iX2, $iY2, False)
			Local $imWidth = $iX2 - $iX1 + 1
			Local $imHeight = $iY2 - $iY1 + 1
			$screenCaptured = True

			Local $winWidDif = $wPos[2] - 400
			Local $winHeiDif = $wPos[3] - 56
			Local $windowWidth = 222

			If $imWidth > $windowWidth Then $windowWidth = $imWidth + $winWidDif;8
			Local $windowHeight = 56 + $imHeight + $winHeiDif;3
			If $windowWidth >= @DesktopWidth Then $windowWidth = @DesktopWidth - $winWidDif
			If $windowHeight >= @DesktopHeight then $windowHeight = @DesktopHeight - $winHeiDif + 3

			Local $childWid = $windowWidth - $winWidDif;8
			Local $childHei = $windowHeight - 56 - $winHeiDif;83

			If $imWidth > $childWid Then
			   If $windowHeight < @DesktopHeight - $winHeiDif Then
				  If $windowHeight + 17 > @DesktopHeight - $winHeiDif Then
					 $windowHeight = @DesktopHeight - $winHeiDif;8
					 $childHei = $windowHeight - 56 - $winHeiDif;83
				  Else
					 $windowHeight += 17
					 $childHei += 17
				  EndIf
			   EndIf
			EndIf

			If $imHeight > $childHei Then
			   If $windowWidth < @DesktopWidth - $winWidDif Then
				  If $windowWidth + 17 > @DesktopWidth - $winWidDif Then
					 $windowWidth = @DesktopWidth - $winWidDif
					 $childWid = $windowWidth - $winWidDif
				  Else
					 $windowWidth += 17
					 $childWid += 17
				  EndIf
			   EndIf
			EndIf

			WinMove($Form2, "", @DesktopWidth / 2 - $windowWidth / 2, @DesktopHeight / 2 - $windowHeight / 2, $windowWidth, $windowHeight)

			Local $child = GUICreate("", $childWid, $childHei, 0, 57, $WS_CHILD, -1, $Form2)
			GUISetBkColor($COLOR_WHITE)

			Local $horizontalSize = -1
			Local $verticalSize = -1

			If $imWidth > $childWid Then $horizontalSize = $imWidth
			If $imHeight > $childHei Then $verticalSize = $imHeight

			_GuiScrollbars_Generate($child, $horizontalSize, $verticalSize, -1, -1, True)

			Local $imageCenter = 0
			If $imWidth < $childWid And $horizontalSize = -1 Then $imageCenter = ($childWid - $imWidth) / 2
			GUICtrlCreatePic($shotPath, $imageCenter, 0, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1)

			GUISetState(@SW_SHOW)
			GUISwitch($Form2)
			GUICtrlSetState($labelMouse, $GUI_HIDE)
			GUICtrlSetState($contButton, $GUI_SHOW)
			GUISetState(@SW_SHOW, $Form2)
		 EndIf
	  EndIf

	  If $nMsg2 == $cancelButton Or $nMsg2 == $GUI_EVENT_CLOSE Then
		 If (FileExists($shotPath)) Then FileDelete($shotPath)
		 GUIDelete($Form2)
		 GUIDelete($Form3)
		 GUISetState(@SW_SHOW, $mainGui)
		 $exit = True
	  EndIf

	  If $nMsg2 == $contButton Then
		 GUIDelete($Form2)
		 Local $result[1]
		 $result[0] = $shotPath
		 ParseFilesToAdd($result)
		 _ArrayAdd($screenShotList, $shotPath)
		 GUISetState(@SW_SHOW, $mainGui)
		 $exit = True
	  EndIf

   WEnd
EndFunc


Func CriticalCheckboxPressed()
	If GUICtrlRead($criticalCheckbox) = $GUI_UNCHECKED Then
		GUICtrlSetBkColor($criticalCheckbox, $CLR_NONE)
	Else
		GUICtrlSetBkColor($criticalCheckbox, $COLOR_YELLOW)
	EndIf
EndFunc


Func SendButtonPressed()
	Local $tempString = ""

	If GUICtrlRead($phoneNumberInput) = "" Then _
		$tempString = $tempString & "Необходимо указать контактный номер телефона" & @CRLF

	If GUICtrlRead($cabinetInput) = "" Then _
		$tempString = $tempString & "Необходимо указать номер кабинета" & @CRLF

	If GUICtrlRead($fullNameInput) = "" Then _
		$tempString = $tempString & "Необходимо указать ФИО" & @CRLF

	If GUICtrlRead($ticketTextEdit) = "" Then _
		$tempString = $tempString & "Необходимо написать текст обращения" & @CRLF

	If $tempString <> "" Then
		MsgBox($MB_ICONWARNING, "", $tempString, 0, $mainGui)
		Return
	EndIf

	Local $needToSendAttachments = True
	If $totalSize > 10000 Then
		Local $answer = MsgBox(BitOr($MB_ICONQUESTION, $MB_YESNO), "Превышен допустимый размер вложений", _
			"Продолжить отправку обращения без вложенных файлов?", 0, $mainGui)
		Switch $answer
			Case $IDYES
				$needToSendAttachments = False
			Case $IDNO
				Return
		EndSwitch
	Else
		SendMessage($needToSendAttachments)
	EndIf
EndFunc


Func FilesAddButtonPressed()
	Local $filesString = FileOpenDialog("Выберите файлы для отправки", _
		Default, "Все(*)", BitOR($FD_FILEMUSTEXIST, $FD_MULTISELECT), "", $mainGui)

	If @error <> 0 Or $filesString = "" Then Return

	Local $resultArray[0]

	$filesString = StringSplit($filesString, "|", $STR_NOCOUNT)
	Local $size = UBound($filesString) - 1

	For $i = 0 To $size
		If Not $size Then
			_ArrayAdd($resultArray, $filesString[0])
			ExitLoop
		EndIf

		If $i Then _ArrayAdd($resultArray, $filesString[0] & "\" & $filesString[$i])
	Next

	ParseFilesToAdd($resultArray)
EndFunc


Func DeleteButtonPressed()
	Local $selected = GUICtrlRead(GUICtrlRead($listView), 2)

	If $selected == 0 Then
		MsgBox($MB_ICONINFORMATION, "", "Необходимо выбрать вложение для удаления", 0, $mainGui)
		Return
	EndIf

	Local $tmp = StringSplit($selected, "|")
	$selected = $tmp[1]
	Local $answer = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO), "", "Вы действительно хотите удалить " & _
		"из вложения файл?" & @CRLF & @CRLF & $selected, 0, $mainGui)
	If $answer = $IDYES Then
		_GUICtrlListView_DeleteItemsSelected($listView)
		GUICtrlSetState($deleteButton, $GUI_DISABLE)
	EndIf

	If _GUICtrlListView_GetItemCount($listView) = 0 Then
		GUICtrlDelete($listView)
		$listView = ""
		GUICtrlDelete($deleteButton)
		$deleteButton = -666
		GUICtrlDelete($attachmentsListGroup)
		$attachmentsListGroup = ""
		Local $lastPos = WinGetPos($mainGui)
		WinMove($mainGui, "", $lastPos[0], $lastPos[1], $dX + 6, $dY + $gapSize * 3)
	Else
		$totalSize = 0
		For $i = 0 To _GUICtrlListView_GetItemCount($listView) - 1
			Local $dataFromListView = _GUICtrlListView_GetItemTextArray($listView, $i)
			$totalSize += $dataFromListView[2]
		Next

		If $totalSize <= 10000 Then
			GUICtrlSetData($attachmentsListGroup, "Список вложений:")
			GUICtrlSetBkColor($attachmentsListGroup, $CLR_NONE)
		EndIf
	EndIf
EndFunc


Func GuiEventDropped()
	If Not IsArray($__aGUIDropFiles) Or Not UBound($__aGUIDropFiles) Then Return

	Local $resultArray[0]
	Local $errorStr = ""

	For $i = 1 To $__aGUIDropFiles[0]
		If StringInStr(FileGetAttrib($__aGUIDropFiles[$i]), "D") Then
			$errorStr &= $__aGUIDropFiles[$i] & @CRLF
		Else
			_ArrayAdd($resultArray, $__aGUIDropFiles[$i])
		EndIf
	Next

	If $errorStr <> "" Then _
		MsgBox($MB_ICONWARNING, "", "Вложениями могут быть только файлы или объекты. " & @CRLF & @CRLF & _
			$errorStr & " - папка(и) и не может быть вложением.", 0, $mainGui)

	If UBound($resultArray) Then ParseFilesToAdd($resultArray)
EndFunc





Func DeleteScreenshots()
   For $i In $screenShotList
	  FileDelete($i)
   Next
EndFunc


Func _MonitorEnumProc($hMonitor, $hDC, $lRect, $lParam)
    Local $Rect = DllStructCreate("int left;int top;int right;int bottom", $lRect)
    $__MonitorList[0][0] += 1
    ReDim $__MonitorList[$__MonitorList[0][0] + 1][5]
    $__MonitorList[$__MonitorList[0][0]][0] = $hMonitor
    $__MonitorList[$__MonitorList[0][0]][1] = DllStructGetData($Rect, "left")
    $__MonitorList[$__MonitorList[0][0]][2] = DllStructGetData($Rect, "top")
    $__MonitorList[$__MonitorList[0][0]][3] = DllStructGetData($Rect, "right")
    $__MonitorList[$__MonitorList[0][0]][4] = DllStructGetData($Rect, "bottom")
    Return 1
EndFunc


Func Mark_Rect()
;   ConsoleWrite("Entering MarkRect" & @CRLF)
   Local $aMouse_Pos, $hMask, $hMaster_Mask, $iTemp
   Local $UserDLL = DllOpen("user32.dll")

   Local $hRectangle_GUI = GUICreate("", $monWid, $monHei, $inX, $inY, $WS_POPUP, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST)
   GUISetBkColor(0x000000)

   $iX1 += Abs($inX)
   While _IsPressed("01", $UserDLL)
	  $aMouse_Pos = MouseGetPos()
	  $aMouse_Pos[0] += Abs($inX)
	  $hMaster_Mask = _WinAPI_CreateRectRgn(0, 0, 0, 0)
;~ 	  $hMaster_Mask = _WinAPI_CreateRectRgn($inX, $inY, $monWid, $monHei)
	  $hMask = _WinAPI_CreateRectRgn($iX1,  $aMouse_Pos[1], $aMouse_Pos[0],  $aMouse_Pos[1] + 1) ; Bottom of rectangle
	  _WinAPI_CombineRgn($hMaster_Mask, $hMask, $hMaster_Mask, 2)
	  _WinAPI_DeleteObject($hMask)
	  $hMask = _WinAPI_CreateRectRgn($iX1, $iY1, $iX1 + 1, $aMouse_Pos[1]) ; Left of rectangle
	  _WinAPI_CombineRgn($hMaster_Mask, $hMask, $hMaster_Mask, 2)
	  _WinAPI_DeleteObject($hMask)
	  $hMask = _WinAPI_CreateRectRgn($iX1 + 1, $iY1 + 1, $aMouse_Pos[0], $iY1) ; Top of rectangle
	  _WinAPI_CombineRgn($hMaster_Mask, $hMask, $hMaster_Mask, 2)
	  _WinAPI_DeleteObject($hMask)
	  $hMask = _WinAPI_CreateRectRgn($aMouse_Pos[0], $iY1, $aMouse_Pos[0] + 1,  $aMouse_Pos[1]) ; Right of rectangle
	  _WinAPI_CombineRgn($hMaster_Mask, $hMask, $hMaster_Mask, 2)
	  _WinAPI_DeleteObject($hMask)
	  ; Set overall region
	  _WinAPI_SetWindowRgn($hRectangle_GUI, $hMaster_Mask)

	  If WinGetState($hRectangle_GUI) < 15 Then GUISetState()
	  Sleep(10)
   WEnd
   $iX1 -= Abs($inX)

   ; Get second mouse position
   $iX2 = $aMouse_Pos[0] - Abs($inX)
   $iY2 = $aMouse_Pos[1]

   ; Set in correct order if required
   If $iX2 < $iX1 Then
	  $iTemp = $iX1
	  $iX1 = $iX2
	  $iX2 = $iTemp
   EndIf

   If $iY2 < $iY1 Then
	  $iTemp = $iY1
	  $iY1 = $iY2
	  $iY2 = $iTemp
   EndIf

   GUIDelete($hRectangle_GUI)
   DllClose($UserDLL)
;   ConsoleWrite("Exiting MarkRect" & @CRLF)
EndFunc   ;==>Mark_Rect


Func ParseFilesToAdd($files)
	If Not IsArray($files) Or Not UBound($files) Then Return

	Local $result[0]
	Local $tooBigFiles[0]

	For $i = 0 To UBound($files) - 1
		Local $driveString = ""
		Local $directoryString = ""
		Local $filenameString = ""
		Local $extensionString = ""
		_PathSplit($files[$i], $driveString, $directoryString, $filenameString, $extensionString)
		Local $fileName = $filenameString & $extensionString
		Local $fileSize = FileGetSize($files[$i]) / 1024

		If $fileSize > 10000 Then
			_ArrayAdd($tooBigFiles, $files[$i])
			_ArrayAdd($tooBigFiles, StringFormat("%.2f", $fileSize))
		Else
			_ArrayAdd($result, $fileName & "|" & StringFormat("%.2f", $fileSize) & "|" & $files[$i], Default, "/////")
		EndIf
	Next

	If UBound($tooBigFiles) Then
		Local $message = "Размер файла(ов)" & @CRLF & @CRLF

		For $i = 0 to UBound($tooBigFiles) - 2 Step 2
			$message &= $tooBigFiles[$i] & "' (" & $tooBigFiles[$i + 1] & " Кб)" & @CRLF
		Next

		$message &= @CRLF & "превышает допустимый размер." & _
			"  Отправка файлов возможна, если их размер не превышает 10 000 Кб." & @CRLF & @CRLF & _
			"Попробуйте заархивировать файл или воспользоваться другим способом передачи" & _
			" (например через сетевую папку или через яндекс/мейл/гугл-диск)"

		MsgBox($MB_ICONWARNING, "Внимание!", $message, 0, $mainGui)
	EndIf

	If Not UBound($result) Then Return

	If $attachmentsListGroup == "" Then
		GUISetState(@SW_LOCK)
		Local $lastPos = WinGetPos($mainGui)

		$currentX = $gapSize * 2
		UpdateCurrentY($ticketTextEdit)
		Local $initialY = $currentY
		$attachmentsListGroup = GUICtrlCreateLabel("Список вложений:", _
			$currentX, $currentY, $dX - $currentX - $gapSize * 2, -1)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)
		GUICtrlSetColor(-1, $textColor)

		UpdateCurrentY($attachmentsListGroup)
		$currentY -= $gapSize

		$listView = GUICtrlCreateListView("Имя файла|Размер (Кб)|Полный путь", _
			$currentX, $currentY, $dX - $currentX - $gapSize * 2, 110)
		_GUICtrlListView_JustifyColumn(-1, 1, 1)
		GUICtrlSetResizing(-1, $GUI_DOCKALL)

		Local $prevCtrlPos = ControlGetPos($mainGui, "", $screenshotButton)
		$deleteButton = GUICtrlCreateButton("Удалить вложение", $prevCtrlPos[0] + $prevCtrlPos[2] + _
			$gapSize * 2, $prevCtrlPos[1], $buttonWidth, $prevCtrlPos[3], $BS_MULTILINE)
		GUICtrlSetState(-1, $GUI_DISABLE)
		GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKWIDTH, $GUI_DOCKHEIGHT))

		UpdateCurrentY($listView)

		$prevCtrlPos = ControlGetPos($mainGui, "", $criticalCheckbox)

		WinMove($mainGui, "", $lastPos[0], $lastPos[1], $dX + 6, $dY + $currentY - $initialY + $gapSize + $prevCtrlPos[3])
	EndIf

	For $str in $result
		GUICtrlCreateListViewItem($str, $listView)
	Next

	_GUICtrlListView_SetColumnWidth($listView, 0, $dX * 0.72)
	_GUICtrlListView_SetColumnWidth($listView, 1, $dX * 0.17)
	_GUICtrlListView_SetColumnWidth($listView, 2, 0)

	$totalSize = 0
	For $i = 0 To _GUICtrlListView_GetItemCount($listView) - 1
		Local $dataFromListView = _GUICtrlListView_GetItemTextArray($listView, $i)
		$totalSize += $dataFromListView[2]
	Next

	If $totalSize > 10000 Then
		GUICtrlSetData($attachmentsListGroup, "Превышен допустимый размер вложенных файлов (10 000 Кб)")
		GUICtrlSetBkColor($attachmentsListGroup, $COLOR_YELLOW)
	EndIf

	GUISetState(@SW_UNLOCK)
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
   $phoneNumber = _AD_GetObjectAttribute(@UserName, "telephoneNumber")
   $mail = _AD_GetObjectAttribute(@UserName, "mail")
   $physicalDeliveryOfficeName = _AD_GetObjectAttribute(@UserName, "physicalDeliveryOfficeName")
   $userPrincipalName = _AD_GetObjectAttribute(@UserName, "userPrincipalName")
   _AD_Close()
EndFunc


Func GetPrinterList()
   For $i = 1 To 1000
	  Local $reg = RegEnumVal( "HKEY_CURRENT_USER\Software\Microsoft\Windows NT\CurrentVersion\Devices", $i)
	  If @error = -1 Then ExitLoop
	  _ArrayAdd($arrPrinterList, $reg)
    Next

    $arrPrinterList[0] = $i - 1
EndFunc




Func SendMessage($needToSendAttachments)
	ProgressOn("", "Идет отправка...", "Пожалуйста, пожождите", -1, -1, $DLG_MOVEABLE)

	Local $message = ""

	If GUICtrlRead($criticalCheckbox) = $GUI_CHECKED Then _
		$message &= ControlGetText($mainGui, "", $criticalCheckbox) & @CRLF & @CRLF

	$message &= GUICtrlRead($ticketTextEdit) & @CRLF

	Local $attachments = ""
	If $needToSendAttachments Then
		If _GUICtrlListView_GetItemCount($listView) > 0 Then $message &= @CRLF & @CRLF & "Список вложений:" & @CRLF

		For $i = 0 To _GUICtrlListView_GetItemCount($listView) - 1
			Local $dataFromListView = _GUICtrlListView_GetItemTextArray($listView, $i)
			$message &= $dataFromListView[1] & " (" & $dataFromListView[2] & " Кб)" & @CRLF
			$attachments &= $dataFromListView[3] & ";"
		Next
	EndIf

	ProgressSet(20)

	$message &= @CRLF & @CRLF & "Инициатор:" & @TAB & @TAB & GUICtrlRead ($fullNameInput) & @CRLF & _
		"Контактный тел.:" & @TAB & GUICtrlRead ($phoneNumberInput) & @CRLF & _
		"Номер кабинета:" & @TAB & GUICtrlRead ($cabinetInput) & @CRLF & @CRLF

	If $city <> "" Then _
		$message &= "Подразделение:" & @TAB & $city & @CRLF

	If $department <> "" Then _
		$message &= "Отдел:" & @TAB & @TAB & @TAB & $department & @CRLF

	If $title <> "" Then _
		$message &= "Должность:" & @TAB & @TAB & $title & @CRLF

	If $mail <> "" Then _
		$message &= "Почта:" & @TAB & @TAB & @TAB & $mail & @CRLF

	If $userPrincipalName <> "" Then _
		$message &= "Учетная запись:" & @TAB & $userPrincipalName & @CRLF

	$message &= @CRLF & "Имя компьютера:" & @TAB & @ComputerName & @CRLF & _
		"Имя пользователя:" & @TAB & @UserName & @CRLF & _
		"Версия ОС:" & @TAB & @TAB & @OSVersion & @CRLF & _
		"Архитектура ОС:" & @TAB & @OSArch & @CRLF & _
		"IP адрес:" & @TAB & @TAB & @IPAddress1 & @CRLF & _
		"Дата обращения:" & @TAB & @YEAR & "." & @MON & "." & @MDAY & "     " & @HOUR & ":" & @MIN

   ProgressSet(30)

   If IsArray($arrPrinterList) Then
	  If UBound($arrPrinterList) > 1 Then
		 $message &= @CRLF & @CRLF & @CRLF & "Список установленных у пользователя принтеров:" & @CRLF
		 For $i = 1 To UBound($arrPrinterList) - 1
			$message = $message & $arrPrinterList[$i] & @CRLF
		 Next
	  EndIf
   EndIf

   ProgressSet(40)

   Local $responce = _INetSmtpMailCom("", "Auto request", "", _
						$emailTo, "Обращение через приложение STP", $message, $attachments, "", "", _
						"", "")
   ProgressSet(100)
   ProgressOff()

   Local $error = @error
   If $responce = 0 And $error = 0 Then
	  Local $tmp = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO), "", _
		 "Ваше обращение успешно отправлено!" & @CRLF & @CRLF & _
		 $exitingMessage & _
		 "Выйти из приложения?", 0, $mainGui)
	  Switch $tmp
	  Case $IDYES
		 DeleteScreenshots()
		 Exit
	  Case $IDNO
	  EndSwitch
   EndIf
EndFunc


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

	 ProgressSet(50)

;~    ConsoleWrite($s_AttachFiles)
    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
;~ 		_ArrayDisplay($S_Files2Attach)
        For $x = 1 To $S_Files2Attach[0] - 1
            $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                $i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
				ConsoleWriteError("file not found")
                SetError(1)
                return 0
            EndIf
        Next
    EndIf
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
   ProgressSet(60)
    ;Authenticated SMTP
    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
   ProgressSet(70)
    ;Update settings
    $objEmail.Configuration.Fields.Update
   ProgressSet(80)
    ; Sent the Message
    $objEmail.Send
	ProgressSet(90)
    if @error then
        SetError(2)
		ProgressOff
        ;return $oMyRet[1]
    EndIf
EndFunc ;==>_INetSmtpMailCom




Func WM_DROPFILES($hWnd, $iMsg, $wParam, $lParam)
	#forceref $hWnd, $lParam
    Switch $iMsg
		Case $WM_DROPFILES
            Local Const $aReturn = _WinAPI_DragQueryFileEx($wParam)
            If UBound($aReturn) Then
                $__aGUIDropFiles = $aReturn
            Else
                Local Const $aError[1] = [0]
                $__aGUIDropFiles = $aError
            EndIf
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc


Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
    Local $tagNMHDR, $event, $hwndFrom, $code
    $tagNMHDR = DllStructCreate("int;int;int", $lParam)
    If @error Then Return
    $event = DllStructGetData($tagNMHDR, 3)

    If $wParam = $listView Then
		If $event = $NM_CLICK Then
			If GUICtrlRead($listView) <> 0 Then
				GUICtrlSetState($deleteButton, $GUI_ENABLE)
			Else
				GUICtrlSetState($deleteButton, $GUI_DISABLE)
			EndIf
		EndIf
	EndIf

    $tagNMHDR = 0
    $event = 0
    $lParam = 0
EndFunc


Func HandleComError()
   If $oMyError.source = "Active Directory" Or $oMyError.source = "" Then Return

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