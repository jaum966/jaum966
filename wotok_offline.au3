#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_UseX64=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****
#include <GUIConstantsEx.au3>
#include <Array.au3>
#include <GuiComboBox.au3>

#NoTrayIcon
#RequireAdmin
Opt("TrayIconHide", 1)

Global $AudienceId, $LCID, $Version, $Product, $Lang, $ExcludedApps, $ProductsToAdd, $ScriptDir
Global $aArraySource[0]
If StringRight(@ScriptDir, 1) = "\" Then
	$ScriptDir = StringTrimRight(@ScriptDir, 1)
Else
	$ScriptDir = @ScriptDir
EndIf
SettingSource()
If UBound($aArraySource) = 0 Then
	MsgBox(0, "Info", "Source files not found.")
	Exit
EndIf

GUICreate("WOTOK 2018.12.27.02 offline", 380, 355)

$guiGroupOnline = GUICtrlCreateGroup("Office 2019", 10, 10, 220, 330)
GUICtrlCreateLabel("Source:", 20, 44, 60, 20)
$guiSource = GUICtrlCreateCombo("", 80, 40, 140, 20)
GUICtrlCreateLabel("Version:", 20, 74, 60, 20)
$guiVersion = GUICtrlCreateCombo("", 80, 70, 140, 20)
GUICtrlCreateLabel("Lang:", 20, 104, 60, 20)
$guiLang = GUICtrlCreateCombo("", 80, 100, 140, 20)
GUICtrlCreateLabel("Channel:", 20, 134, 60, 20)
$guiChannel = GUICtrlCreateCombo("", 80, 130, 140, 20)
GUICtrlCreateLabel("Product:", 20, 164, 60, 20)
$guiProduct = GUICtrlCreateCombo("", 80, 160, 140, 20)
GUICtrlCreateLabel("Display:", 20, 194, 60, 20)
$guiDisplay = GUICtrlCreateCombo("", 80, 190, 140, 20)
GUICtrlCreateLabel("Telemetry:", 20, 224, 60, 20)
$guiTelemetry = GUICtrlCreateCombo("", 80, 220, 140, 20)
$guiInstallOffice = GUICtrlCreateButton("Install Office", 20, 300, 200, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("ExcludedApps", 240, 10, 130, 220)
$idCBword       = GUICtrlCreateCheckbox("word", 250, 30, 100, 25)
$idCBexcel      = GUICtrlCreateCheckbox("excel", 250, 50, 100, 25)
$idCBoutlook    = GUICtrlCreateCheckbox("outlook", 250, 70, 100, 25)
$idCBpowerpoint = GUICtrlCreateCheckbox("powerpoint", 250, 90, 100, 25)
$idCBaccess     = GUICtrlCreateCheckbox("access", 250, 110, 100, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBpublisher  = GUICtrlCreateCheckbox("publisher", 250, 130, 100, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBonenote    = GUICtrlCreateCheckbox("onenote", 250, 150, 100, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBonedrive   = GUICtrlCreateCheckbox("onedrive", 250, 170, 100, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
$idCBlync       = GUICtrlCreateCheckbox("lync", 250, 190, 100, 25)
GUICtrlSetState(-1, $GUI_CHECKED)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlCreateGroup("Install", 240, 240, 130, 100)
$idCBvisio        = GUICtrlCreateCheckbox("VisioPro2019", 250, 260, 100, 25)
$idCBproject      = GUICtrlCreateCheckbox("ProjectPro2019", 250, 280, 100, 25)
$idCBseco         = GUICtrlCreateCheckbox("KMS Seco", 250, 300, 100, 25)
GUICtrlCreateGroup("", -99, -99, 1, 1)

GUICtrlSetData($guiProduct, "ProPlus2019Volume|ProPlus2019Retail","ProPlus2019Volume")
GUICtrlSetData($guiChannel, "Dogfood::DevMain|Dogfood::CC|Microsoft::DevMain|Microsoft::CC|Microsoft::LTSC|Insiders::DevMain|Insiders::CC|Production::CC|Production::LTSC","Production::CC")
GUICtrlSetData($guiDisplay, "True|False", "True")
GUICtrlSetData($guiTelemetry, "Set Disable|Do not set", "Set Disable")

_KMSsecoVisibility()
GUICtrlSetData($guiSource, _ArrayToString($aArraySource, "|"), $aArraySource[0])
SettingVersion()
ComboLang()
SettingChannel()
GUISetState(@SW_SHOW)


While 1
	Switch GUIGetMsg()
		Case $GUI_EVENT_CLOSE
			ExitLoop
        Case $guiInstallOffice
			InstallOffice()
		Case $guiProduct
			_KMSsecoVisibility()
		Case $guiSource
			SettingVersion()
			ComboLang()
		Case $guiVersion
			ComboLang()
	EndSwitch
WEnd

Func SettingSource()
	CheckSourceDir($ScriptDir, 1)
	Local $aArrayDrive = DriveGetDrive("ALL")
	For $i = 1 To $aArrayDrive[0]
		CheckSourceDir($aArrayDrive[$i], 0)
	Next
EndFunc

Func CheckSourceDir($vSourceDir, $scd)
	If FileExists ($vSourceDir & "\Office\Data") Then
		Local $hSearch = FileFindFirstFile($vSourceDir & "\Office\Data\*.*")
		If $hSearch <> -1 Then
			Local $sFileName = ""
			While 1
				$sFileName = FileFindNextFile($hSearch)
				If @error Then ExitLoop
				If StringInStr(FileGetAttrib($vSourceDir & "\Office\Data\" & $sFileName), "D") > 0 Then
					If FileExists($vSourceDir & "\Office\Data\" & $sFileName & "\stream.x64.x-none.dat") Then
						If $vSourceDir = $ScriptDir And $scd = 1 Then
							_ArrayAdd($aArraySource, "ScriptDir")
						Else
							_ArrayAdd($aArraySource, $vSourceDir)
						EndIf
						ExitLoop
					EndIf
				EndIf
			WEnd
			FileClose($hSearch)
		EndIf
	EndIf
EndFunc

Func SettingVersion()
	$SourceDir = GUICtrlRead($guiSource)
	If $SourceDir = "ScriptDir" Then $SourceDir = $ScriptDir
	_GUICtrlComboBox_ResetContent($guiVersion)
	Local $hSearch = FileFindFirstFile($SourceDir & "\Office\Data\*.*")
	If $hSearch <> -1 Then
		Local $sFileName = ""
		While 1
			$sFileName = FileFindNextFile($hSearch)
			If @error Then ExitLoop
			If StringInStr(FileGetAttrib($SourceDir & "\Office\Data\" & $sFileName), "D") > 0 Then
				If FileExists($SourceDir & "\Office\Data\" & $sFileName & "\stream.x64.x-none.dat") Then
					GUICtrlSetData($guiVersion, $sFileName, $sFileName)
					$Version = $sFileName
				EndIf
			EndIf
		WEnd
		FileClose($hSearch)
	EndIf
EndFunc

Func SettingChannel()
	$AudienceData = GUICtrlRead($guiChannel)
	If $AudienceData = "Dogfood::DevMain"    Then $AudienceId = "EA4A4090-DE26-49D7-93C1-91BFF9E53FC3"
	If $AudienceData = "Dogfood::CC"         Then $AudienceId = "F3260CF1-A92C-4C75-B02E-D64C0A86A968"
	If $AudienceData = "Microsoft::DevMain"  Then $AudienceId = "B61285DD-D9F7-41F2-9757-8F61CBA4E9C8"
	If $AudienceData = "Microsoft::CC"       Then $AudienceId = "5462EEE5-1E97-495B-9370-853CD873BB07"
	If $AudienceData = "Microsoft::LTSC"     Then $AudienceId = "1D2D2EA6-1680-4C56-AC58-A441C8C24FF9"
	If $AudienceData = "Insiders::DevMain"   Then $AudienceId = "5440FD1F-7ECB-4221-8110-145EFAA6372F"
	If $AudienceData = "Insiders::CC"        Then $AudienceId = "64256AFE-F5D9-4F86-8936-8840A6A4F5BE"
	If $AudienceData = "Production::CC"      Then $AudienceId = "492350F6-3A01-4F97-B9C0-C7C6DDF67D60"
	If $AudienceData = "Production::LTSC"    Then $AudienceId = "F2E724C1-748F-4B47-8FB8-8E0D210E9208"
EndFunc

Func ComboLang()
	_GUICtrlComboBox_ResetContent($guiLang)
	$SourceDir = GUICtrlRead($guiSource)
	If $SourceDir = "ScriptDir" Then $SourceDir = $ScriptDir
	$Version = GUICtrlRead($guiVersion)
	Local $aArrayLang[0]
	Local $hSearch = FileFindFirstFile($SourceDir & "\Office\Data\"&$Version&"\stream.x64.*.dat")
	If $hSearch <> -1 Then
		Local $sFileName = ""
		While 1
			$sFileName = FileFindNextFile($hSearch)
			If @error Then ExitLoop
			If $sFileName = "stream.x64.x-none.dat" Then ContinueLoop
			If StringRight ($sFileName, 10) = ".proof.dat" Then ContinueLoop
			$sFileName = StringReplace($sFileName, "stream.x64.", "")
			$sFileName = StringReplace($sFileName, ".dat", "")
			_ArrayAdd($aArrayLang, $sFileName)
		WEnd
		FileClose($hSearch)
	EndIf
	GUICtrlSetData($guiLang, _ArrayToString($aArrayLang, "|"), $aArrayLang[0])
EndFunc

Func SettingLang()
	$Lang = GUICtrlRead($guiLang)
	If $Lang = "ar-sa" Then $LCID = "1025"
	If $Lang = "bg-bg" Then $LCID = "1026"
	If $Lang = "cs-cz" Then $LCID = "1029"
	If $Lang = "en-us" Then $LCID = "1033"
	If $Lang = "es-es" Then $LCID = "3082"
	If $Lang = "et-ee" Then $LCID = "1061"
	If $Lang = "fi-fi" Then $LCID = "1035"
	If $Lang = "fr-fr" Then $LCID = "1036"
	If $Lang = "he-il" Then $LCID = "1037"
	If $Lang = "hi-in" Then $LCID = "1081"
	If $Lang = "hr-hr" Then $LCID = "1050"
	If $Lang = "hu-hu" Then $LCID = "1038"
	If $Lang = "it-it" Then $LCID = "1040"
	If $Lang = "ja-jp" Then $LCID = "1041"
	If $Lang = "ko-kr" Then $LCID = "1042"
	If $Lang = "lt-lt" Then $LCID = "1063"
	If $Lang = "lv-lv" Then $LCID = "1062"
	If $Lang = "ms-my" Then $LCID = "1086"
	If $Lang = "nb-no" Then $LCID = "1044"
	If $Lang = "nl-nl" Then $LCID = "1043"
	If $Lang = "pl-pl" Then $LCID = "1045"
	If $Lang = "pt-br" Then $LCID = "1046"
	If $Lang = "pt-pt" Then $LCID = "2070"
	If $Lang = "ro-ro" Then $LCID = "1048"
	If $Lang = "ru-ru" Then $LCID = "1049"
	If $Lang = "sk-sk" Then $LCID = "1051"
	If $Lang = "sl-si" Then $LCID = "1060"
	If $Lang = "sv-se" Then $LCID = "1053"
	If $Lang = "th-th" Then $LCID = "1054"
	If $Lang = "tr-tr" Then $LCID = "1055"
	If $Lang = "uk-ua" Then $LCID = "1058"
	If $Lang = "vi-vn" Then $LCID = "1066"
	If $Lang = "zh-cn" Then $LCID = "2052"
	If $Lang = "zh-tw" Then $LCID = "1028"
	If $Lang = "sr-latn-rs" Then $LCID = "9242"
EndFunc

Func InstallOffice()
	GUICtrlSetState($guiInstallOffice, $GUI_DISABLE)
	sleep(500)
	$Display = GUICtrlRead($guiDisplay)
	If _IsChecked($idCBseco) Then _KMSsecoInstall()
	$SourceDir = GUICtrlRead($guiSource)
	If $SourceDir = "ScriptDir" Then $SourceDir = $ScriptDir
	$Version = GUICtrlRead($guiVersion)
	SettingLang()
	SettingChannel()
	SettingProduct()
	SettingExcludedApps()
	If Not FileExists(@CommonFilesDir & "\microsoft shared\ClickToRun\OfficeClickToRun.exe") Then
		UnPack($SourceDir & "\Office\Data\"&$Version&"\i640.cab",@CommonFilesDir & "\microsoft shared\ClickToRun")
		UnPack($SourceDir & "\Office\Data\"&$Version&"\i64"&$LCID&".cab",@CommonFilesDir & "\microsoft shared\ClickToRun")
	EndIf
	RunWait (@CommonFilesDir & "\microsoft shared\ClickToRun\OfficeClickToRun.exe" & _
		" deliverymechanism="&$AudienceId & _
		" platform=x64" & _
		" culture="&$Lang & _
		" displaylevel=" & $Display & _
		" cdnbaseurl.16=http://officecdn.microsoft.com/pr/"&$AudienceId & _
		" baseurl.16="""&$SourceDir&"""" & _
		" version.16="&$Version & _
		" productstoadd="&$ProductsToAdd & _
		$ExcludedApps)
	$Telemetry = GUICtrlRead($guiTelemetry)
	If $Telemetry = "Set Disable" Then _DisableTelemetry()
	GUICtrlSetState($guiInstallOffice, $GUI_ENABLE)
EndFunc

Func SettingProduct()
	$Product = GUICtrlRead($guiProduct)
	$ProductsToAdd = $Product&".16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBvisio) Then $ProductsToAdd &= "|VisioPro2019Volume.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBvisio) Then $ProductsToAdd &= "|VisioPro2019Retail.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBproject) Then $ProductsToAdd &= "|ProjectPro2019Volume.16_"&$Lang&"_x-none"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBproject) Then $ProductsToAdd &= "|ProjectPro2019Retail.16_"&$Lang&"_x-none"
EndFunc

Func UnPack($sUnPackFileName, $sUnPackDestination)
	$oUnPackFSO = ObjCreate("Scripting.FileSystemObject")
	If Not $oUnPackFSO.FolderExists($sUnPackDestination) Then
		$oUnPackFSO.CreateFolder($sUnPackDestination)
	EndIf
	$WshShell = ObjCreate("Shell.Application")
	With $WshShell
		.NameSpace($sUnPackDestination).Copyhere (.NameSpace($sUnPackFileName).Items)
	EndWith
EndFunc

Func SettingExcludedApps()
	$ExcludedApps = ""
	Local $aArrayExcludedApps[0]
	If _IsChecked($idCBword)       Then _ArrayAdd($aArrayExcludedApps, "word")
	If _IsChecked($idCBexcel)      Then _ArrayAdd($aArrayExcludedApps, "excel")
	If _IsChecked($idCBoutlook)    Then _ArrayAdd($aArrayExcludedApps, "outlook")
	If _IsChecked($idCBpowerpoint) Then _ArrayAdd($aArrayExcludedApps, "powerpoint")
	If _IsChecked($idCBaccess)     Then _ArrayAdd($aArrayExcludedApps, "access")
	If _IsChecked($idCBpublisher)  Then _ArrayAdd($aArrayExcludedApps, "publisher")
	If _IsChecked($idCBonenote)    Then _ArrayAdd($aArrayExcludedApps, "onenote")
	If _IsChecked($idCBonedrive)   Then _ArrayAdd($aArrayExcludedApps, "onedrive")
	                                    _ArrayAdd($aArrayExcludedApps, "groove")
	If _IsChecked($idCBlync)       Then _ArrayAdd($aArrayExcludedApps, "lync")
	If UBound($aArrayExcludedApps) Then $ExcludedApps = " " & $Product & ".excludedapps.16=" & _ArrayToString($aArrayExcludedApps, ",")
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBvisio) And _IsChecked($idCBonedrive) Then $ExcludedApps &= " VisioPro2019Volume.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBvisio) And _IsChecked($idCBonedrive) Then $ExcludedApps &= " VisioPro2019Retail.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Volume" And _IsChecked($idCBproject) And _IsChecked($idCBonedrive) Then $ExcludedApps &= " ProjectPro2019Volume.excludedapps.16=onedrive,groove"
	If $Product = "ProPlus2019Retail" And _IsChecked($idCBproject) And _IsChecked($idCBonedrive) Then $ExcludedApps &= " ProjectPro2019Retail.excludedapps.16=onedrive,groove"
EndFunc

Func _IsChecked($idControlID)
    Return BitAND(GUICtrlRead($idControlID), $GUI_CHECKED) = $GUI_CHECKED
EndFunc

Func _KMSsecoVisibility()
	If Not FileExists (@SystemDir & "\SppExtComObjHook.dll") And FileExists($ScriptDir & "\64\SppExtComObjHook.dll") And GUICtrlRead($guiProduct) = "ProPlus2019Volume" Then
		GUICtrlSetState($idCBseco, $GUI_SHOW + $GUI_CHECKED)
	Else
		GUICtrlSetState($idCBseco, $GUI_HIDE + $GUI_UNCHECKED)
	EndIf
EndFunc

Func _KMSsecoInstall()
	FileCopy ($ScriptDir & "\64\SppExtComObjHook.dll", @SystemDir, 1)
	RegDelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\55c92734-d682-4d71-983e-d6ec3f16059f")
	RegDelete("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\SoftwareProtectionPlatform\0ff1ce15-a989-479d-af46-f275c6370663")
	RegWrite("HKLM\SOFTWARE\Policies\Microsoft\Windows NT\CurrentVersion\Software Protection Platform", "NoGenTicket", "REG_DWORD", 1)
	RegWrite("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Image File Execution Options\SppExtComObj.exe", "Debugger", "REG_SZ", "rundll32.exe SppExtComObjHook.dll,PatcherMain")
	$oWMIService = ObjGet("winmgmts:\\.\root\cimv2")
	If IsObj($oWMIService) Then
		$oCollection = $oWMIService.ExecQuery("SELECT Version FROM SoftwareLicensingService")
		If IsObj($oCollection) Then
			For $oItem In $oCollection
				$oItem.SetKeyManagementServiceMachine("172.16.16.16")
				$oItem.SetKeyManagementServicePort("1688")
			Next
		EndIf
	EndIf
EndFunc

Func _DisableTelemetry()
	RegWrite("HKLM\Software\Policies\Microsoft\Office\16.0\osm", "Enablelogging", "REG_DWORD", 0)
	RegWrite("HKLM\Software\Policies\Microsoft\Office\16.0\osm", "EnableUpload", "REG_DWORD", 0)
	RegWrite("HKLM\Software\Microsoft\Office\Common\ClientTelemetry", "DisableTelemetry", "REG_DWORD", 1)
	RegWrite("HKCU\Software\Policies\Microsoft\Office\16.0\osm", "Enablelogging", "REG_DWORD", 0)
	RegWrite("HKCU\Software\Policies\Microsoft\Office\16.0\osm", "EnableUpload", "REG_DWORD", 0)
	RegWrite("HKCU\Software\Microsoft\Office\Common\ClientTelemetry", "DisableTelemetry", "REG_DWORD", 1)
EndFunc