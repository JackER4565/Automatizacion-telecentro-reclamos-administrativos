#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GUIListView.au3>
#include <WindowsConstants.au3>
#include <D:\UDF\wd_core.au3> ;https://github.com/Danp2/au3WebDriver TY Danp2
#include <D:\UDF\wd_helper.au3> ;https://github.com/Danp2/au3WebDriver TY Danp2
#include <Crypt.au3>
#include <array.au3>
#include <Date.au3>
#include <_INetSmtpMailCom.au3> ;included - https://github.com/groubis/typebtomail/blob/master/_INetSmtpMailCom.au3 TY Jos


Local $sDesiredCapabilities, $sSession,$inpt_clnt, $sElement, $sButton,$sValue, $sUser, $sPass, $chrome_handle, $hndl
Local $aElements, $iLines
global $oldra_datoscorp = 0, $oldra_datosres = 0, $oldra_telcorp = 0, $oldra_telres = 0, $oldra_trama = 0, $body_aux, $primera_vuelta = 0

			_WD_UpdateDriver('chrome')

#Region chrome


Call(SetupChrome)
_WD_Startup()


If @error <> $_WD_ERROR_Success Then
	Exit -1
EndIf

#EndRegion chrome

#Region INI
If Not FileExists(@ScriptDir & "\config.ini") Then
	IniWrite(@ScriptDir & "\config.ini", "config", "Version", "0.0")
EndIf
If IniRead(@ScriptDir & "\config.ini", "config", "UsernameCRM", "0") == 0 Then
	Do
		$user = InputBox("Usuario", "Aca es donde deberias poner el usuario de CRM.","", "", "", "135")
	Until Not @error
	IniWrite(@ScriptDir & "\config.ini", "config", "UsernameCRM", $user)
EndIf

If IniRead(@ScriptDir & "\config.ini", "config", "NoeslaPasswordCRM", "0") == 0 Then
	Do
		$Pass = InputBox("Password", "Aca es donde deberias poner la contraseña de CRM.", "", "?", "", "135")
	Until Not @error
			Local $dEncrypted = StringEncrypt(True, $Pass, 'securepassword')
			IniWrite(@ScriptDir & "\config.ini", "config", "NoeslaPasswordCRM", $dEncrypted)
EndIf

$dEncrypted = IniRead(@ScriptDir & "\config.ini", "config", "NoeslaPasswordCRM", "0")
Local $sDecrypted = StringEncrypt(False, $dEncrypted, 'securepassword')
Global $Pass = $sDecrypted
$user = IniRead(@ScriptDir & "\config.ini", "config", "UsernameCRM", "0")
Sleep(1000)

#EndRegion INI

#Region ### START Koda GUI section ### Form=
Global $gui_glo = GUICreate("RaS cRm:", 341, 200, 294, 149)

Global $listado = GUICtrlCreateListView("Cola|Cantidad", 8, 8, 325, 150, BitOR($GUI_SS_DEFAULT_LISTVIEW, $WS_VSCROLL))
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 0, 240)
GUICtrlSendMsg(-1, $LVM_SETCOLUMNWIDTH, 1, 75)
_GUICtrlListView_JustifyColumn(GUICtrlGetHandle($listado), 0, 2)
_GUICtrlListView_JustifyColumn(GUICtrlGetHandle($listado), 1, 2)


$boton_go = GUICtrlCreateButton("G0", 8, 160, 90, 25)
$boton_clean = GUICtrlCreateButton("Limpiar", 100, 160, 90, 25)
$boton_pass = GUICtrlCreateButton("Cambiar Contraseña", 192, 160, 141, 25)

GUISetState(@SW_SHOW)
#EndRegion ### END Koda GUI section ###

Global $timer = TimerInit()

While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
		Case $GUI_EVENT_CLOSE
			_WD_DeleteSession($sSession)
			_WD_Shutdown()
			Exit

		Case $boton_pass
			Do
				$user = InputBox("Usuario", "Aca es donde deberias poner el usuario de CRM.","", "", "", "135")
			Until Not @error
				IniWrite(@ScriptDir & "\config.ini", "config", "UsernameCRM", $user)
			Do
				$Pass = InputBox("Password", "Aca es donde deberias poner la contraseña de CRM.", "", "?", "", "135")
			Until Not @error
				Local $dEncrypted = StringEncrypt(True, $Pass, 'securepassword')
				IniWrite(@ScriptDir & "\config.ini", "config", "NoeslaPasswordCRM", $dEncrypted)
		Case $boton_clean
			_GUICtrlListView_DeleteAllItems($listado)
		Case $boton_go
			GOOO()
		Case Else
			$fDiff = TimerDiff($timer)
			;(tiempo pa q corra)	$timeaux = ($fDiff - 300000) / 1000
			;	$timeaux = StringFormat("%d", $timeaux)
			WinSetTitle($gui_glo, "", "RaS cRm: " & $fDiff)
			If $fDiff >= 300000 Then

				GOOO()
				$timer = TimerInit()
			EndIf
			;--- m4
	EndSwitch
WEnd

Func GOOO()
	WinSetTitle($gui_glo, "", "RaS cRm: " & "Go")
				_GUICtrlListView_DeleteAllItems($listado)
			$body_aux = ""

			#Region PRIMERA PARTE
			WinSetTitle($gui_glo, "", "RaS cRm: " & "Creando Session")
			$sSession = _WD_CreateSession($sDesiredCapabilities)
			_WD_LoadWait($sSession)

			; NAVEGA A USUARIOS ###
			_WD_Navigate($sSession, "http://" & $user & ":" & $Pass &"@crm.telecentro.local")

			; Locate elements (inputs)
			$sPass = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/form/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/table/tbody/tr[6]/td[2]/div/input')

			; clear inputs
			_WD_ElementAction($sSession, $sPass, 'clear')

			; Set element's contents (inputs y opt)
			_WD_ElementAction($sSession, $sPass, 'value', $Pass)
			Sleep(500)

			; Click login button
			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/form/table/tbody/tr[2]/td/table/tbody/tr[3]/td/table/tbody/tr[1]/td[2]/table/tbody/tr[7]/td[2]/table/tbody/tr[1]/td/table/tbody/tr/td/nobr/input')
			_WD_ElementAction($sSession, $sButton, 'click')
			_WD_LoadWait($sSession, 200)

;~
			; Navega a crm

			_WD_Navigate($sSession, "http://crm.telecentro.local//Cliente/ReclamoAdministrativo/ReclamoCierre.aspx?SubMenu=475")
			_WD_LoadWait($sSession, 200)
			_WD_Navigate($sSession, "http://crm.telecentro.local//Cliente/ReclamoAdministrativo/ReclamoCierre.aspx?SubMenu=475")
			_WD_LoadWait($sSession, 200)

			$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/form/div[6]/table/tbody/tr/td/table/tbody/tr/td/div/iframe")
			_WD_FrameEnter($sSession, $sElement)


			$hoy = _NowDate()
			Sleep(1000)
				$sElement = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[3]/div[2]/input')
			_WD_ElementAction($sSession, $sElement, 'value', $hoy)
			Sleep(500)
			WinSetTitle($gui_glo, "", "RaS cRm: " & "Fin Primera parte")
			#EndRegion PRIMERA PARTE



		#Region ESC CNOC/CEOP - DATOS CORP
			WinSetTitle($gui_glo, "", "RaS cRm: " & "ESC CNOC/CEOP - DATOS CORP")
			_WD_ElementOptionSelect($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/select/option[58]")

				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(500)

			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[4]/div[3]/button')
			_WD_ElementAction($sSession, $sButton, 'click')
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(1000)

			$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr", "", True)     ; Retrieve the number of table rows
			if _WD_ElementAction($sSession, $aElements[0], "attribute", "class") = "b-table-empty-row" Then
				$iLines = 0
			Else
				$iLines = UBound($aElements)
			EndIf
			if $primera_vuelta = 0 Then $oldra_datoscorp = $iLines
			if $oldra_datoscorp < $iLines Then $body_aux = "Cola: ESC CNOC/CEOP - DATOS CORP -- RAs nuevos: " & $iLines & @CRLF
			$oldra_datoscorp = $iLines
			$index = _GUICtrlListView_AddItem($listado, "ESC CNOC/CEOP - DATOS CORP ")
			_GUICtrlListView_AddSubItem($listado, $index, $iLines, 1)
			$iLines = 0
			Sleep(1000)
		#EndRegion ESC CNOC/CEOP - DATOS CORP
		#Region ESC CNOC/CEOP - DATOS RES
			WinSetTitle($gui_glo, "", "RaS cRm: " & "ESC CNOC/CEOP - DATOS RES")
			_WD_ElementOptionSelect($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/select/option[59]")
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(500)
			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[4]/div[3]/button')
			_WD_ElementAction($sSession, $sButton, 'click')
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(1000)

			$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr", "", True)     ; Retrieve the number of table rows
			if _WD_ElementAction($sSession, $aElements[0], "attribute", "class") = "b-table-empty-row" Then
				$iLines = 0
			Else
				$iLines = UBound($aElements)
			EndIf
						if $primera_vuelta = 0 Then $oldra_datosres = $iLines
			if $oldra_datosres < $iLines Then $body_aux &= "Cola: ESC CNOC/CEOP - DATOS RES -- RAs nuevos: " & $iLines & @CRLF
			 $oldra_datosres = $iLines
			$index = _GUICtrlListView_AddItem($listado, "ESC CNOC/CEOP - DATOS RES ")
			_GUICtrlListView_AddSubItem($listado, $index, $iLines, 1)
			$iLines = 0
			Sleep(1000)
		#EndRegion ESC CNOC/CEOP - DATOS RES
		#Region ESC CNOC/CEOP - TELEF RES
			WinSetTitle($gui_glo, "", "RaS cRm: " & "ESC CNOC/CEOP - TELEF RES")
			_WD_ElementOptionSelect($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/select/option[60]")
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(500)
			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[4]/div[3]/button')
			_WD_ElementAction($sSession, $sButton, 'click')
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(1000)

			$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr", "", True)     ; Retrieve the number of table rows
			if _WD_ElementAction($sSession, $aElements[0], "attribute", "class") = "b-table-empty-row" Then
				$iLines = 0
			Else
				$iLines = UBound($aElements)
			EndIf
			if $primera_vuelta = 0 Then $oldra_telres = $iLines
			if $oldra_telres < $iLines Then $body_aux &= "Cola: ESC CNOC/CEOP - TELEF RES -- RAs nuevos: " & $iLines & @CRLF
			$oldra_telres = $iLines
			$index = _GUICtrlListView_AddItem($listado, "ESC CNOC/CEOP - TELEF RES ")
			_GUICtrlListView_AddSubItem($listado, $index, $iLines, 1)
			$iLines = 0
			Sleep(1000)
		#EndRegion ESC CNOC/CEOP - TELEF RES
		#Region ESC CNOC/CEOP - TELEF CORP
			WinSetTitle($gui_glo, "", "RaS cRm: " & "ESC CNOC/CEOP - TELEF CORP")
			_WD_ElementOptionSelect($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/select/option[61]")
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(500)
			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[4]/div[3]/button')
			_WD_ElementAction($sSession, $sButton, 'click')
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(1000)

			$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr", "", True)     ; Retrieve the number of table rows
			if _WD_ElementAction($sSession, $aElements[0], "attribute", "class") = "b-table-empty-row" Then
				$iLines = 0
			Else
				$iLines = UBound($aElements)
			EndIf
			if $primera_vuelta = 0 Then $oldra_telcorp = $iLines
			if $oldra_telcorp < $iLines Then $body_aux &= "Cola: ESC CNOC/CEOP - TELEF CORP -- RAs nuevos: " & $iLines & @CRLF
			$oldra_telcorp = $iLines
			$index = _GUICtrlListView_AddItem($listado, "ESC CNOC/CEOP - TELEF CORP ")
			_GUICtrlListView_AddSubItem($listado, $index, $iLines, 1)
			$iLines = 0
			Sleep(1000)
		#EndRegion ESC CNOC/CEOP - TELEF CORP
		#Region TRAMA TELEFONICA
			WinSetTitle($gui_glo, "", "RaS cRm: " & "TRAMA TELEFONICA")
			_WD_ElementOptionSelect($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[1]/div[2]/select/option[139]")
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(500)
			$sButton = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[3]/div[1]/div/div/div[1]/div/div[4]/div[3]/button')
			_WD_ElementAction($sSession, $sButton, 'click')
				$cargando = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, '/html/body/div/div[4]')
				Do
					$respuesta = _WD_ElementAction($sSession, $cargando, "attribute", "style")
				Sleep(100)
				Until $respuesta = "display: none;"
			Sleep(1000)

			$aElements = _WD_FindElement($sSession, $_WD_LOCATOR_ByXPath, "/html/body/div/div[3]/div[2]/div/div/div/div/div[1]/div/div/div[2]/div/table/tbody/tr", "", True)     ; Retrieve the number of table rows
			if _WD_ElementAction($sSession, $aElements[0], "attribute", "class") = "b-table-empty-row" Then
				$iLines = 0
			Else
				$iLines = UBound($aElements)
			EndIf
			if $primera_vuelta = 0 Then $oldra_trama = $iLines
			if $oldra_trama < $iLines Then $body_aux &= "Cola: TRAMA TELEFONICA -- RAs nuevos: " & $iLines & @CRLF
			$oldra_trama = $iLines
			$index = _GUICtrlListView_AddItem($listado, "TRAMA TELEFONICA")
			_GUICtrlListView_AddSubItem($listado, $index, $iLines, 1)
			$iLines = 0
			Sleep(1000)
		#EndRegion TRAMA TELEFONICA

		if not($body_aux == "") Then
			WinSetTitle($gui_glo, "", "RaS cRm: " & "Enviando correo")
			sendmail($body_aux)
			$body_aux = ""
		EndIf
		$primera_vuelta = 1
		WinSetTitle($gui_glo, "", "RaS cRm: " & "Fin")
		_WD_DeleteSession($sSession)
EndFunc

func sendmail($Body)
	$rc = _INetSmtpMailCom($SmtpServer, $FromName, $FromAddress, $ToAddress, $Subject, $Body, $AttachFiles, $CcAddress, $BccAddress, $Importance, $Username, $Pass, $IPPort, $ssl)
EndFunc

Func SetupChrome()
;~ 	$_WD_DEBUG = $_WD_DEBUG_Error
$_WD_DEBUG = $_WD_DEBUG_None
_WD_Option('Driver', @ScriptDir & '\chromedriver.exe')
_WD_Option('Port', 9515)
;~ _WD_Option('DriverParams', '--log-path="' & @ScriptDir & '\chromeRA.log"')

$sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"goog:chromeOptions": {"w3c": true }}}}'
;~ $sDesiredCapabilities = '{"capabilities": {"alwaysMatch": {"goog:chromeOptions": {"w3c": true, "args": ["--headless", "--allow-running-insecure-content"] }}}}'
EndFunc


Func StringEncrypt($bEncrypt, $sData, $sPassword)
	_Crypt_Startup() ; Start the Crypt library.
	Local $vReturn = ''
	If $bEncrypt Then ; If the flag is set to True then encrypt, otherwise decrypt.
		$vReturn = _Crypt_EncryptData($sData, $sPassword, $CALG_RC4)
	Else
		$vReturn = BinaryToString(_Crypt_DecryptData($sData, $sPassword, $CALG_RC4))
	EndIf
	_Crypt_Shutdown() ; Shutdown the Crypt library.
	Return $vReturn
EndFunc   ;==>StringEncrypt
