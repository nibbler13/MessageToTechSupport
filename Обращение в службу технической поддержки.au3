#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Icon=icon.ico
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#pragma compile(ProductVersion, 2.1)
#pragma compile(UPX, true)
#pragma compile(CompanyName, 'ООО Клиника ЛМС')
#pragma compile(FileDescription, Программа для создания и отправки обращения в службу технической поддержки пользователей)
#pragma compile(LegalCopyright, )
#pragma compile(ProductName, Обращение в службу технической поддержки)

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

#Region ### Fields ###
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
#EndRegion ### End of fields

ProgressOn("Запуск программы", "Сбор данных о пользователе", "Пожалуйста, немного подождите", -1, -1, BitOR($DLG_NOTONTOP, $DLG_MOVEABLE))
ReadUserInfo()

#Region ### START Koda GUI section ### Form=
$Form1 = GUICreate("", 600, 450, -1, -1, Default, $WS_EX_ACCEPTFILES)
GUICtrlCreateGroup("", -10, -10, 610, 460)
;~ GUICtrlSetState(-1, $GUI_DROPACCEPTED)

GUISetFont(10, $FW_NORMAL , 0, "Arial", $Form1, $CLEARTYPE_QUALITY)

$Label1 = GUICtrlCreateLabel("Обращение в службу технической поддержки клиники Будь Здоров", 10, 10, 580, 20, $SS_CENTER)
GUICtrlSetFont(-1, 11, $FW_SEMIBOLD)
GUICtrlSetColor(-1, "0x3030AA")
$Label5 = GUICtrlCreateLabel("внутренний номер: 603    для регионов: 30-494    городской: (495) 782-88-82", 10, 30, 580, 20, $SS_CENTER)
GUICtrlSetColor(-1, $COLOR_GRAY)

$Group1 = GUICtrlCreateGroup("Укажите информацию о себе:", 10, 70, 580, 124)
GUICtrlSetFont(-1, 10,  $FW_SEMIBOLD)
GUICtrlSetResizing(-1, $GUI_DOCKALL)

$Label2 = GUICtrlCreateLabel("Контактный номер телефона:", 20, 103, 180, 20, $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
$Input1 = GUICtrlCreateInput("", 210, 100, 363, 24)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
If $telephoneNumber <> "" Then
   GUICtrlSetData(-1, $telephoneNumber)
EndIf

$Label3 = GUICtrlCreateLabel("Номер кабинета:", 20, 133, 180, 20, $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
$Input2 = GUICtrlCreateInput("", 210, 130, 363, 24)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
If $physicalDeliveryOfficeName <> "" Then
   GUICtrlSetData(-1, $physicalDeliveryOfficeName)
EndIf

$Label4 = GUICtrlCreateLabel("ФИО:", 20, 163, 180, 20, $SS_RIGHT)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
$Input3 = GUICtrlCreateInput("", 210, 160, 363, 24)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
If $userName <> "" Then
   GUICtrlSetData(-1, $userName)
Else
   GUICtrlSetData(-1, $user)
EndIf

GUICtrlCreateGroup("", -99, -99, 1, 1)

$Group2 = GUICtrlCreateGroup("Текст вашего обращения:", 10, 200, 580, 190)
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetFont(-1, 10,  $FW_SEMIBOLD)
$Edit1 = GUICtrlCreateEdit("", 20, 230, 560, 150, BitOR($ES_AUTOVSCROLL,$ES_WANTRETURN,$WS_VSCROLL))
GUICtrlSetResizing(-1, $GUI_DOCKALL)
GUICtrlSetState(-1, $GUI_FOCUS)
GUICtrlCreateGroup("", -99, -99, 1, 1)

$ButtonAddFiles = GUICtrlCreateButton("Вложить файл", 9, 400, 82, 40, $BS_MULTILINE)
GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKWIDTH, $GUI_DOCKHEIGHT))
GUICtrlSetFont(-1, Default, $FW_LIGHT)

Local $prevButPos = ControlGetPos($Form1, "", $ButtonAddFiles)
$ButtonScreenshot = GUICtrlCreateButton("Сделать cнимок экрана", $prevButPos[0] + $prevButPos[2] + 13, _
$prevButPos[1], 130, $prevButPos[3], $BS_MULTILINE)
GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKWIDTH, $GUI_DOCKHEIGHT))
GUICtrlSetFont(-1, Default, $FW_LIGHT)

$ButtonSend = GUICtrlCreateButton("Отправить обращение", 411, 400, 180, 40)
GUICtrlSetResizing(-1, BitOr($GUI_DOCKBOTTOM, $GUI_DOCKWIDTH, $GUI_DOCKHEIGHT))
GUICtrlSetFont(-1, Default, $FW_HEAVY)
ProgressSet(100)
ProgressOff()
GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

GetPrinterList()
GUIRegisterMsg($WM_NOTIFY, "MY_WM_NOTIFY")
GUIRegisterMsg($WM_DROPFILES, 'WM_DROPFILES')

While 1
   $nMsg = GUIGetMsg()
   Switch $nMsg
	  Case $GUI_EVENT_CLOSE
		 DeleteScreenshots()
		 Exit
	  Case $ButtonSend
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
			Local $needToSendAttachments = True
;~ 			   ConsoleWrite($totalSize)
			If $totalSize / 1024 > 10000 Then
			   Local $answer = MsgBox(BitOr($MB_ICONQUESTION, $MB_YESNO), "Превышен допустимый размер вложений", _
				  "Продолжить отправку обращения без вложенных файлов?")
			   Switch $answer
			   Case $IDYES
				  $needToSendAttachments = False
			   Case $IDNO
				  ContinueLoop
			   EndSwitch
			Else
			EndIf
			SendMessage($needToSendAttachments)
		 EndIf
	  Case $ButtonScreenshot
		 CreateScreenshotWindow()
	  Case $ButtonAddFiles
		 Local $filesString = FileOpenDialog("Выберите файлы для отправки", _
			Default, "Все(*)", BitOR($FD_FILEMUSTEXIST, $FD_MULTISELECT), "", $Form1)

		 If @error <> 0 Or $filesString = "" Then
			ContinueLoop
		 EndIf

		 ParseFilesToAdd($filesString)
	  Case $deleteButton
		 Local $selected = GUICtrlRead(GUICtrlRead($listView), 2)
		 If $selected == 0 Then
			MsgBox($MB_ICONINFORMATION, "", "Необходимо выбрать вложение для удаления")
		 Else
			Local $tmp = StringSplit($selected, "|")
			$selected = $tmp[1]
			Local $answer = MsgBox(BitOR($MB_ICONQUESTION, $MB_YESNO), "", "Вы действительно хотите удалить " & _
			   "из вложения файл?" & @CRLF & @CRLF & $selected)
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
			   Local $lastPos = WinGetPos($Form1)
			   WinMove($Form1, "", $lastPos[0], $lastPos[1], 606, 478)
			Else
			   $totalSize = 0
			   For $i = 0 To _GUICtrlListView_GetItemCount($listView) - 1
				  Local $dataFromListView = _GUICtrlListView_GetItemTextArray($listView, $i)
				  $totalSize += FileGetSize($dataFromListView[3])
			   Next

			   If $totalSize / 1024 <= 10000 Then
				  GUICtrlSetData($attachmentsListGroup, "Список вложений:")
				  GUICtrlSetBkColor($attachmentsListGroup, $CLR_NONE)
			   EndIf
			EndIf
		 EndIf
	  Case $GUI_EVENT_DROPPED
		 If UBound($__aGUIDropFiles) == 2 Then
			If StringInStr(FileGetAttrib($__aGUIDropFiles[1]), "D") Then
			   MsgBox($MB_ICONWARNING, "", "Вложениями могут быть только файлы или объекты. " & $__aGUIDropFiles[1] & _
				  " - папка и не может быть вложением.")
			Else
			ParseFilesToAdd($__aGUIDropFiles[1])
		 EndIf

		 Else
			Local $resultStr = ""
			Local $errorStr = ""
			For $i = 1 To $__aGUIDropFiles[0]
			   If StringInStr(FileGetAttrib($__aGUIDropFiles[$i]), "D") Then
				  $errorStr &= $__aGUIDropFiles[$i] & @CRLF
			   Else
				  Local $driveString = ""
				  Local $directoryString = ""
				  Local $filenameString = ""
				  Local $extensionString = ""
				  $tmp = _PathSplit($__aGUIDropFiles[$i], $driveString, $directoryString, $filenameString, $extensionString)
;~ 				  $tmp = _PathSplit($__aGUIDropFiles[$i], "", "", "", "")
				  If $resultStr == "" Then $resultStr &= $tmp[1] & $tmp[2]
				  $resultStr &= "|" & $tmp[3] & $tmp[4]
			   EndIf
			Next

			If $errorStr <> "" Then
			   MsgBox($MB_ICONWARNING, "", "Вложениями могут быть только файлы или объекты. " & @CRLF & _
				  $errorStr & " - папка(и) и не может быть вложением.")
			EndIf

			If $resultStr <> "" Then ParseFilesToAdd($resultStr)
		 EndIf
   EndSwitch
WEnd


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
    Return 1 ; Return 1 to continue enumeration
EndFunc   ;==>_MonitorEnumProc


Func CreateScreenshotWindow()
   GUISetState(@SW_HIDE, $Form1)

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

   Local $Form2 = GUICreate("Создание снимка экрана", 400, 56, -1, -1, -1, $WS_EX_TOOLWINDOW + $WS_EX_TOPMOST, $Form1)
   GUISetFont(10, $FW_NORMAL , 0, "Arial", $Form2, $CLEARTYPE_QUALITY)
   Local $cancelButton = GUICtrlCreateButton("Отмена", 8, 8, 78, 40)
   GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKLEFT, $GUI_DOCKSIZE))
   Local $contButton = GUICtrlCreateButton("Продолжить", 278, 8, 114, 40)
   GUICtrlSetResizing(-1, BitOR($GUI_DOCKTOP, $GUI_DOCKRIGHT, $GUI_DOCKSIZE))
   GUICtrlSetFont(-1, Default, $FW_HEAVY)
   GUICtrlSetState(-1, $GUI_HIDE)

   Local $labelMouse = GUICtrlCreateLabel("Выделите мышью нужную область экрана", _
	  86, 19, 314, 40, $SS_CENTER)
   GUICtrlSetFont(-1, Default, $FW_HEAVY)

   GUISetState(@SW_SHOW, $Form2)

   Local $exit = False
   Local $screenCaptured = False
   Local $shotPath

   While Not $exit
	  $nMsg2 = GUIGetMsg()

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

;~ 			$tmp = WinGetPos($Form2)
;~ 			_ArrayDisplay($tmp)


			Local $winWidDif = $wPos[2] - 400
			Local $winHeiDif = $wPos[3] - 56
			Local $windowWidth = 222
;~ 			ConsoleWrite("DIF: " & $winWidDif & " " & $winHeiDif & @CRLF)
;~ 			ConsoleWrite("IM: " & $imWidth & " " & $imHeight & @CRLF)
			If $imWidth > $windowWidth Then $windowWidth = $imWidth + $winWidDif;8
			Local $windowHeight = 56 + $imHeight + $winHeiDif;3
			If $windowWidth >= @DesktopWidth Then $windowWidth = @DesktopWidth - $winWidDif
			If $windowHeight >= @DesktopHeight then $windowHeight = @DesktopHeight - $winHeiDif + 3

			Local $childWid = $windowWidth - $winWidDif;8
			Local $childHei = $windowHeight - 56 - $winHeiDif;83

;~ 			ConsoleWrite("WIN: " & $windowWidth & " " & $windowHeight & @CRLF)
;~ 			ConsoleWrite("CH: " & $childWid & " " & $childHei & @CRLF)

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

;~ 			Local $tmp1 = WinGetClientSize($Form2)
;~ 			ConsoleWrite("CL: " & $tmp1[0] & " " & $tmp1[1] & @CRLF)


			Local $child = GUICreate("", $childWid, $childHei, 0, 57, $WS_CHILD, -1, $Form2)
			GUISetBkColor($COLOR_WHITE)
;			ConsoleWrite("imcenter: " & $imageCenter & @CRLF)
			Local $horizontalSize = -1
			Local $verticalSize = -1
;~ 			ConsoleWrite($childWid & " " & $imWidth & " " & $childHei & " " & $imHeight & @CRLF)
;~ 			ConsoleWrite($windowWidth & " " & $windowHeight & @CRLF)
			If $imWidth > $childWid Then $horizontalSize = $imWidth
			If $imHeight > $childHei Then $verticalSize = $imHeight
;~ 			If $horizontalSize > -1 Then $verticalSize += 1
;~ 			If $verticalSize > -1 Then $horizontalSize += 1
;~ 			ConsoleWrite($horizontalSize & " " & $verticalSize & @CRLF)

			_GuiScrollbars_Generate($child, $horizontalSize, $verticalSize, -1, -1, True)


			Local $imageCenter = 0
			If $imWidth < $childWid And $horizontalSize = -1 Then $imageCenter = ($childWid - $imWidth) / 2
			GUICtrlCreatePic($shotPath, $imageCenter, 0, $iX2 - $iX1 + 1, $iY2 - $iY1 + 1)

			GUISetState(@SW_SHOW)
			GUISwitch($Form2)
			GUICtrlSetState($labelMouse, $GUI_HIDE)
			GUICtrlSetState($contButton, $GUI_SHOW)
			GUISetState(@SW_SHOW, $Form2)
			;ContinueLoop
		 EndIf
	  EndIf

	  If $nMsg2 == $cancelButton Or $nMsg2 == $GUI_EVENT_CLOSE Then
		 If (FileExists($shotPath)) Then FileDelete($shotPath)
		 GUIDelete($Form2)
		 GUIDelete($Form3)
		 GUISetState(@SW_SHOW, $Form1)
		 $exit = True
	  EndIf

	  If $nMsg2 == $contButton Then
		 GUIDelete($Form2)
		 ParseFilesToAdd($shotPath)
		 _ArrayAdd($screenShotList, $shotPath)
		 GUISetState(@SW_SHOW, $Form1)
		 $exit = True
	  EndIf

   WEnd
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
   Local $choosenFiles = StringSplit($files, "|")
   Local $result[0]
   If IsArray($choosenFiles) Then
			If UBound($choosenFiles) = 2 Then
				Local $driveString = ""
				Local $directoryString = ""
				Local $filenameString = ""
				Local $extensionString = ""
				Local $fileName = _PathSplit($choosenFiles[1], $driveString, $directoryString, $filenameString, $extensionString)
;~ 			   Local $fileName = _PathSplit($choosenFiles[1], "", "", "", "")
			   Local $fileSize = FileGetSize($choosenFiles[1]) / 1024
			   If $fileSize > 10000 Then
				  MsgBox($MB_ICONWARNING, "Внимание!", "Размер файла" & @CRLF & @CRLF & $fileName[3] & $fileName[4] & _
					 "' (" & StringFormat("%.2f", $fileSize) & " Кб)" & @CRLF & @CRLF & _
					 "превышает допустимый размер." & _
					 " Отправка файлов возможна, если их размер не превышает 10 000 Кб." & @CRLF & @CRLF & _
					 "Попробуйте заархивировать файл или воспользоваться другим способом передачи" & _
					 " (например через сетевую папку или через яндекс/мейл/гугл-диск")
			   Else
				  _ArrayAdd($result, $fileName[3] & $fileName[4] & "|" & _
					 StringFormat("%.2f", $fileSize) & "|" & $choosenFiles[1], Default, "/////")
			   EndIf
			EndIf

			Local $tooBigFiles[0]

			For $i = 2 To UBound($choosenFiles) - 1
			   Local $fileName = $choosenFiles[1] & "\" & $choosenFiles[$i]
			   Local $fileSize = FileGetSize($fileName) / 1024
			   If $fileSize > 10000 Then
				  _ArrayAdd($tooBigFiles, $choosenFiles[$i])
				  _ArrayAdd($tooBigFiles, StringFormat("%.2f", $fileSize))
			   Else
				  _ArrayAdd($result, $choosenFiles[$i] & "|" & StringFormat("%.2f", $fileSize) & _
					 "|" & $fileName, Default, "/////")
			   EndIf
			Next

			If UBound($tooBigFiles) > 0 Then
			   Local $message = "Размер файла(ов)" & @CRLF & @CRLF

			   For $i = 0 to UBound($tooBigFiles) - 2 Step 2
				  $message &= $tooBigFiles[$i] & "' (" & $tooBigFiles[$i + 1] & " Кб)" & @CRLF
			   Next

			   $message &= @CRLF & "превышает допустимый размер." & _
					 "  Отправка файлов возможна, если их размер не превышает 10 000 Кб." & @CRLF & @CRLF & _
					 "Попробуйте заархивировать файл или воспользоваться другим способом передачи" & _
					 " (например через сетевую папку или через яндекс/мейл/гугл-диск)"

			   MsgBox($MB_ICONWARNING, "Внимание!", $message)
			EndIf
		 EndIf

		 $choosenFiles = ""

		 If UBound($result) == 0 Then
			Return
		 EndIf

		 If $attachmentsListGroup == "" Then
			Local $lastPos = WinGetPos($Form1)
			WinMove($Form1, "", $lastPos[0], $lastPos[1], 606, 633)

			$attachmentsListGroup = GUICtrlCreateGroup("Список вложений:", 10, 396, 580, 150)
			GUICtrlSetFont(-1, 10,  $FW_SEMIBOLD)
			GUICtrlSetResizing(-1, $GUI_DOCKALL)

			$listView = GUICtrlCreateListView("Имя файла|Размер (Кб)|Полный путь", 20, 426, 560, 110)
			_GUICtrlListView_SetColumnWidth(-1, 0, 445)
			_GUICtrlListView_SetColumnWidth(-1, 2, 0)
			_GUICtrlListView_JustifyColumn(-1, 1, 1)
			GUICtrlSetResizing(-1, $GUI_DOCKALL)

			Local $prevButtonPos = ControlGetPos($Form1, "", $ButtonScreenshot)
			$deleteButton = GUICtrlCreateButton("Удалить выбранное вложение", $prevButtonPos[0] + $prevButtonPos[2] + _
			13, $prevButtonPos[1], 151, $prevButtonPos[3], $BS_MULTILINE)
			GUICtrlSetState(-1, $GUI_DISABLE)
		 EndIf

		 For $str in $result
			GUICtrlCreateListViewItem($str, $listView)
		 Next
		 _GUICtrlListView_SetColumnWidth($listView, 2, 0)

		 $totalSize = 0
		 For $i = 0 To _GUICtrlListView_GetItemCount($listView) - 1
			Local $dataFromListView = _GUICtrlListView_GetItemTextArray($listView, $i)
			$totalSize += FileGetSize($dataFromListView[3])
		 Next

		 If $totalSize / 1024 > 10000 Then
			GUICtrlSetData($attachmentsListGroup, "Список вложений: Превышен допустимый размер вложенных файлов (10 000 Кб)")
			GUICtrlSetBkColor($attachmentsListGroup, $COLOR_YELLOW)
		 EndIf
EndFunc


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

	ConsoleWrite('@@ Debug(' & @ScriptLineNumber & ') : $GUI_RUNDEFMSG = ' & $GUI_RUNDEFMSG & @CRLF & '>Error code: ' & @error & @CRLF) ;### Debug Console

    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_DROPFILES


Func MY_WM_NOTIFY($hWnd, $Msg, $wParam, $lParam)
;~     #forceref $hWndGUI, $MsgID, $wParam
    Local $tagNMHDR, $event, $hwndFrom, $code
    $tagNMHDR = DllStructCreate("int;int;int", $lParam);NMHDR (hwndFrom, idFrom, code)
    If @error Then Return
    $event = DllStructGetData($tagNMHDR, 3)
    Select
	   Case $wParam = $listView
		   Select
			   Case $event = $NM_CLICK
				  If GUICtrlRead($listView) <> 0 Then
					 GUICtrlSetState($deleteButton, $GUI_ENABLE)
				  Else
					 GUICtrlSetState($deleteButton, $GUI_DISABLE)
				  EndIf
			EndSelect
	   EndSelect
    $tagNMHDR = 0
    $event = 0
    $lParam = 0
EndFunc


Func SendMessage($needToSendAttachments)
	ProgressOn("", "Идет отправка.", "Пожалуйста, немного пожождите.", -1, -1, $DLG_MOVEABLE)
	Local $message = GUICtrlRead($Edit1) & @CRLF

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

	$message &= @CRLF & @CRLF & "Инициатор:" & @TAB & @TAB & GUICtrlRead ($Input3) & @CRLF & _
					 "Контактный тел.:" & @TAB & GUICtrlRead ($Input1) & @CRLF & _
					 "Номер кабинета:" & @TAB & GUICtrlRead ($Input2) & @CRLF & @CRLF

	If $city <> "" Then _
	  $message = $message & "Подразделение:" & @TAB & $city & @CRLF

	If $department <> "" Then _
	  $message = $message & "Отдел:" & @TAB & @TAB & @TAB & $department & @CRLF

	If $title <> "" Then _
	  $message = $message & "Должность:" & @TAB & @TAB & $title & @CRLF

	If $mail <> "" Then _
	  $message = $message & "Почта:" & @TAB & @TAB & @TAB & $mail & @CRLF

	If $userPrincipalName <> "" Then _
	  $message = $message & "Учетная запись:" & @TAB & $userPrincipalName & @CRLF

	$message = $message & @CRLF & "Имя компьютера:" & @TAB & @ComputerName & @CRLF & _
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
						"", "Обращение через приложение STP", $message, $attachments, "", "", _
						"", "")
	ProgressSet(100)
	ProgressOff()

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
			DeleteScreenshots()
			Exit
		Case $IDNO
		EndSwitch
	EndIf
EndFunc


Func ReadUserInfo()
   ProgressSet(10)
   _AD_Open()

   If @error Then
	  Return
   EndIf

   $userName = _AD_GetObjectAttribute(@UserName, "displayName")
   ProgressSet(20)
   $city = _AD_GetObjectAttribute(@UserName, "l")
   ProgressSet(30)
   $department = _AD_GetObjectAttribute(@UserName, "department")
   ProgressSet(40)
   $title = _AD_GetObjectAttribute(@UserName, "title")
   ProgressSet(50)
   $telephoneNumber = _AD_GetObjectAttribute(@UserName, "telephoneNumber")
   ProgressSet(60)
   $mail = _AD_GetObjectAttribute(@UserName, "mail")
   ProgressSet(70)
   $physicalDeliveryOfficeName = _AD_GetObjectAttribute(@UserName, "physicalDeliveryOfficeName")
   ProgressSet(80)
   $userPrincipalName = _AD_GetObjectAttribute(@UserName, "userPrincipalName")
   ProgressSet(90)
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


Func _INetSmtpMailCom($s_SmtpServer, $s_FromName, $s_FromAddress, $s_ToAddress, $s_Subject = "", $as_Body = "", _
	$s_AttachFiles = "", $s_CcAddress = "", $s_BccAddress = "", $s_Username = "", $s_Password = "",$IPPort=25, $ssl=0)
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

    If $s_AttachFiles <> "" Then
        Local $S_Files2Attach = StringSplit($s_AttachFiles, ";")
        For $x = 1 To $S_Files2Attach[0] - 1
            $S_Files2Attach[$x] = _PathFull ($S_Files2Attach[$x])
            If FileExists($S_Files2Attach[$x]) Then
                $objEmail.AddAttachment ($S_Files2Attach[$x])
            Else
                $i_Error_desciption = $i_Error_desciption & @lf & 'File not found to attach: ' & $S_Files2Attach[$x]
				ConsoleWriteError("file not found")
                SetError(1)
                return -1
            EndIf
        Next
    EndIf

    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver") = $s_SmtpServer
    $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = $IPPort
	ProgressSet(60)

    If $s_Username <> "" Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = $s_Username
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = $s_Password
    EndIf
    If $Ssl Then
        $objEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    EndIf
	ProgressSet(70)

    $objEmail.Configuration.Fields.Update
	ProgressSet(80)

    $objEmail.Send
    if @error then
        SetError(2)
		ProgressOff()
		Return -1
    EndIf
EndFunc


Func HandleComError()
	If $oMyError.source = "Active Directory" Then
		ConsoleWrite("source is empty")
		Return
	EndIf

	ProgressOff()
	Msgbox($MB_ICONERROR, "Обращение в техподдержку", "Возникла ошибка при отправке." & @CRLF & _
			"Обратитесь в техподдержку по телефону: 603 или 30-494"    & @CRLF  & @CRLF & _
			"Описание ошибки: " & @TAB & $oMyError.description)

Endfunc