With WScript
Set objFS=.CreateObject("Scripting.FileSystemObject")
Set objWshShell=.CreateObject("WScript.Shell")
Set objMessage=.CreateObject("Scripting.Dictionary")
End With
debugMode=True
If Not debugMode Then On Error Resume Next
Const ERROR_SUCCESS=0
Const ERROR_EBP_BASE=&HB1050000
Const EBP_BADPARAM=&H0001
Const EBP_NOTSUPPORTOS=&H0002
Const EBP_NOTADMIN=&H0003
Const EBP_ALREADYEXEC=&H0004
Const EBP_CANNOTEXEC=&H0005
Const EBP_NOTSUPPORTPC=&H0006
Const EBP_NOTCOMPATIBLEPC=&H0007
Const EBP_NOTCOMPATIIMAGE=&H0008
Const EBP_CANNOTDELFOLDER=&H0009
Const EBP_UPTODATE=&H000A
Const EBP_EXECCANCELED=&H0100
Const EBP_NOPERMISSION=&H0101
Const EBP_BADDEVDRIVER=&H0102
Const EBP_INCOMPATIIMAGE=&H0103
Const EBP_CANNOTREADIMAGE=&H0104
Const EBP_CORRUPTEDIMAGE=&H0105
Const EBP_BADENCPASFORMAT=&H0106
Const EBP_SOMEERROR_BASE=&H0100
Const EBP_ESPERROR_BASE=&H1000
Const EBP_EXTRACTCANCELED=&H0200
Const EBP_CANNOTEXTRACT=&H0201
Const MsgBoxWidth=600
Const MsgBoxHeight=400
Const MsgBoxWidth2=610
Const MsgBoxHeight2=420
Const FdBoxWidth=450
Const FdBoxHeight=250
Const strMutexName="TOS_BIOS_Package_2"
Const strPopupTitle="TOSHIBA BIOS Package %v"
Const strToshibaFolder="Toshiba"
Const strPackageFolder="BIOS_Package"
Const strCloseFlgName="_closeexecutingbar_.flg"
Const strBaseFolder="C:\TOSHBIOS."
Const strDefFolderExt="UPD"
Const strAutoFolderExt="AUTO"
Const strFdFolderExt="FD"
Const strSmsFolderExt="SMS"
Const strBiosFolderName="BIOS"
Const strEcFolderName="EC"
Const strUndefinedFolder="Undefined_Folder"
Const strExe32Name="NCHGBIOS2.EXE"
Const strExe64Name="NCHGBIOS2x64.EXE"
Const strChgBiosName="CHGBIOSF.EXE"
Const strExeESP32Name="NCHGBIOSESP.EXE"
Const strExeESP64Name="NCHGBIOSESPx64.EXE"
Const strChgBiosEFIName="CHGBIOSA.EFI"
Const strExe32Name3="NCHGBIOS3.EXE"
Const strExe64Name3="NCHGBIOS3x64.EXE"
Const strChgBiosPegasus10PName="TOSFIRMUP.EFI"
Const strChgBiosPegasus10PToolName="startup.nsh"
Const strBiosFilePrefix="bio"
Const strBiosFileExt="t.com"
Const strBiosPegasus10PFileExt="t.bin"
Const strBiosPegasus10PToolName="H2OFFT-Sx64.efi"
Const strEcFilePrefix="ec"
Const strEcFileExt="bin"
Const strEcPegasus10PFileExt="txt"
Const strEcPegasus10PToolName="ChgPgEC.efi"
Const strMachineIdPegasus10P="00A1"
Const strEcTypePegasus10P="ECPG"
Const strRequiredBiosVersionPegasus10P="1.60"
Const strPlatformReader32Name="PlatformReader.exe"
Const strPlatformReader64Name="PlatformReaderx64.exe"
Const PLATFORM_FAILED=0
Const PLATFORM_UNKNOWN=1
Const PLATFORM_CRESCENTBAY_OR_OLDER=2
Const PLATFORM_SKYLAKE_OR_LATER_CSM=3
Const PLATFORM_SKYLAKE_OR_LATER_UEFI=4
isCrescentBayOrOlder=False
isSkylakeOrLaterCSM=False
isSkylakeOrLaterUEFI=False
isPegasus10P=False
isESPErrorScheme=False
Const strPostExeName="TosDspProfileConvert.exe"
Const strSilentParam="/silent"
Const strAutoParam="/auto"
Const strFdParam="/fd"
Const strSmsParam="/sms"
Const strAnyOSParam="/anyos"
Const strAllowReWriteParam="/force"
Const strAllowVersionDownParam="/allowversiondown"
Const strNoCleanParam="/noclean"
Const strNoRebootParam="/noreboot"
Const strWarnTypeSParam="/ws"
Const strLangParam="/lang"
Const strSvPassParam="/svpass"
Const strOwnerParam="/owner"
Const strExecParam="/exec"
Const strSuppressPopupParam="/suppresspopup"
Const strCheckOnlyParam="/checkonly"
Const strDebugMode="/debugMode"
Const strParamFileExt=".prm"
Const strParamOwnerSection="owner"
Const strParamX64Key="x64"
Const strParamAdminKey="admin"
Const strParamLangIDKey="langid"
Const strEmbedFileExt=".emb"
Const strEmbedSvPassKey="svpass"
Const strEmbedSilentKey="silent"
Const strEmbedNoRebootKey="noreboot"
Const strEmbedAllowReWriteKey="force"
Const strDefLangId="0409"
Const strIniName="version.ini"
Const strIniPackSection="BIOS_Package"
Const strIniVerKey="version"
specialCase=Array(Array(1,"0075","EC71","<","V1.50"),Array(1,"0076","EC72","<","V1.50"),Array(1,"0077","EC74","<","V1.20"),Array(1,"007D","EC77","<","V1.50"),Array(1,"007E","EC78","<","V1.30"),Array(1,"007F","EC7A","<","V1.10"),Array(1,"0080","EC79","<","V1.10"),Array(1,"0081","EC7C","<","V1.20"),Array(0))
exitCode=ERROR_SUCCESS
ownerMode=False
cleanMode=True
rebootMode=False
noRebootMode=False
silentMode=False
autoMode=False
fdMode=False
smsMode=False
anyOSMode=False
checkOnlyMode=False
allowReWriteMode=False
allowVerDownMode=False
WarnTypeS=False
x64Mode=False
suppressPopup=False
isAdmin=False
ExecutingBarClosed=False
LangID=strDefLangId
targetFolder=""
SvPass=""
ErrorMsg=""
Const MinimumBatteryLifePercent=30
needBattChk=True
FONT_COLOR_WHITE=vbWhite
FONT_COLOR_BLUE=vbBlue
FONT_COLOR_RED=vbRed
FONT_COLOR_GREEN=RGB(0,128,0)
FONT_COLOR_LIME=vbGreen
FONT_COLOR_MAROON=RGB(128,0,0)
FONT_COLOR_ORANGE=RGB(238,120,0)
myPath=WScript.ScriptFullName
With objFS
myFolder=.GetParentFolderName(myPath)
myName=.GetFileName(myPath)
paramFile=.BuildPath(myFolder,.GetBaseName(myName)&strParamFileExt)
embedFile=.BuildPath(myFolder,.GetBaseName(myName)&strEmbedFileExt)
iniFile=.BuildPath(myFolder,strIniName)
End With
curFolder=objWshShell.CurrentDirectory
OSType=GetOSType
If objFS.FileExists(paramFile)Then
x64Mode=(GetIniKeyVal(paramFile,strParamOwnerSection,strParamX64Key)="1")
isAdmin=(GetIniKeyVal(paramFile,strParamOwnerSection,strParamAdminKey)="1")
LangID=GetIniKeyVal(paramFile,strParamOwnerSection,strParamLangIDKey)
End If
If Len(LangID)<>4 Then LangID=FormatToHex(Hex(GetLocale),4)
If objFS.FileExists(embedFile)Then
If GetIniKeyVal(embedFile,"",strEmbedSilentKey)="1" Then silentMode=True
If GetIniKeyVal(embedFile,"",strEmbedNoRebootKey)="1" Then noRebootMode=True
If GetIniKeyVal(embedFile,"",strEmbedAllowReWriteKey)="1" Then allowReWriteMode=True
SvPass=GetIniKeyVal(embedFile,"",strEmbedSvPassKey)
If(SvPass<>"")And Not debugMode Then objFS.DeleteFile embedFile,True
End If
PackageVer=GetIniKeyVal(iniFile,strIniPackSection,strIniVerKey)
PopupTitle=StringReplace(strPopupTitle,"%v",PackageVer)
If WScript.Arguments.Count>0 Then
For i=0 To WScript.Arguments.Count-1
arg=WScript.Arguments(i)
LCarg=LCase(arg)
If arg=strDebugMode Then
debugMode=True
PopupMsg"Debug mode."
ElseIf LCarg=strSilentParam Then
silentMode=True
If debugMode Then PopupMsg"Silent mode."
ElseIf LCarg=strAutoParam Then
If targetFolder="" Then
autoMode=True
If debugMode Then PopupMsg"Auto mode."
Else
exitCode=EBP_BADPARAM
End If
ElseIf LCarg=strFdParam Then
fdMode=True
If debugMode Then PopupMsg"FD mode."
ElseIf LCarg=strSmsParam Then
smsMode=True
If debugMode Then PopupMsg"SMS mode."
ElseIf LCarg=strAnyOSParam Then
anyOSMode=True
If debugMode Then PopupMsg"Any OS mode."
ElseIf LCarg=strAllowReWriteParam Then
allowReWriteMode=True
If debugMode Then PopupMsg"Allow Re-Write mode."
ElseIf LCarg=strAllowVersionDownParam Then
allowVerDownMode=True
If debugMode Then PopupMsg"Allow Version Down mode."
ElseIf LCarg=strNoCleanParam Then
cleanMode=False
If debugMode Then PopupMsg"No Clean mode."
ElseIf LCarg=strNoRebootParam Then
noRebootMode=True
If debugMode Then PopupMsg"No Reboot mode."
ElseIf LCarg=strWarnTypeSParam Then
WarnTypeS=True
If debugMode Then PopupMsg"Warning type: S"
ElseIf LCarg=strSuppressPopupParam Then
suppressPopup=True
If debugMode Then PopupMsg"Suppress Popup mode."
ElseIf InStr(LCarg,strSvPassParam&"=")=1 Then
SvPass=Mid(arg,Len(strSvPassParam)+2)
If SvPass="" Then exitCode=EBP_BADPARAM
If debugMode Then PopupMsg"SvPass = "&SvPass
ElseIf InStr(LCarg,strLangParam&"=")=1 Then
If Len(LCarg)=(Len(strLangParam)+4)Then
LCarg=Right(LCarg,3)
If LCarg="enu" Then
LangID="0409"
ElseIf LCarg="jpn" Then
LangID="0411"
Else
exitCode=EBP_BADPARAM
End If
If debugMode Then PopupMsg"LangID = "&LangID
Else
exitCode=EBP_BADPARAM
End If
ElseIf InStr(LCarg,strOwnerParam&"=")=1 Then
If Len(LCarg)>(Len(strOwnerParam)+1)Then
ownerMode=True
If debugMode Then PopupMsg"Owner mode."
Else
exitCode=EBP_BADPARAM
End If
ElseIf(Left(LCarg,1)<>"/")and(targetFolder="")Then
targetFolder=arg
End If
If exitCode<>ERROR_SUCCESS Then Exit For
Next
End If
If exitCode=ERROR_SUCCESS Then
If silentMode Then
If fdMode Or autoMode Or smsMode Or(targetFolder<>"")Then
exitCode=EBP_BADPARAM
Else
autoMode=True
suppressPopup=True
End If
ElseIf autoMode Then
If fdMode Or smsMode Then
exitCode=EBP_BADPARAM
Else
If targetFolder<>"" Then
If Instr(UCase(targetFolder),UCase(strBaseFolder))=1 Then
strExt=UCase(Mid(targetFolder,Len(strBaseFolder)+1))
Else
strExt=""
End If
If debugMode Then PopupMsg"strExt = """&strExt&""""
If strExt=UCase(strSmsFolderExt)Then
noRebootMode=True
smsMode=True
If debugMode Then PopupMsg"Changed to SMS mode."
ElseIf strExt=UCase(strFdFolderExt)Then
autoMode=False
fdMode=True
targetFolder=""
If debugMode Then PopupMsg"Changed to FD mode."
ElseIf strExt<>UCase(strAutoFolderExt)Then
exitCode=EBP_BADPARAM
End If
End If
End If
ElseIf smsMode Then
If fdMode Or(targetFolder<>"")Then
exitCode=EBP_BADPARAM
Else
autoMode=True
noRebootMode=True
End If
ElseIf fdMode Then
If suppressPopup Then exitCode=EBP_BADPARAM
End If
If allowVerDownMode And Not allowReWriteMode Then exitCode=EBP_BADPARAM
End If
If Not autoMode Then suppressPopup=False
If debugMode Then On Error Goto 0
TranslateMessages
If cleanMode And Not ownerMode Then
cleanMode=False
If debugMode Then PopupMsg"Changed to No Clean mode."
End If
PopupTitle=ReplMessage("CAPTIONTITLE","%v",PackageVer)
If debugMode Then PopupMsg"myPath = "&myPath&vbCrLf&"myFolder = "&myFolder&vbCrLf&"myName = "&myName&vbCrLf&"paramFile = "&paramFile&vbCrLf&"iniFile = "&iniFile&vbCrLf&"curFolder = "&curFolder&vbCrLf&vbCrLf&"OS Type = "&OSType&vbCrLf&"isAdmin = "&isAdmin&vbCrLf&"LangID = """&LangID&""""&vbCrLf&"PackageVer = """&PackageVer&""""&vbCrLf&"targetFolder = """&targetFolder&""""&vbCrLf&"PopupTitle = """&PopupTitle&""""
Do
If exitCode<>ERROR_SUCCESS Then
ErrorMsg=GetMessage("BADPARAM")
Exit Do
End If
If OSType<OS_WinXP Then
exitCode=EBP_NOTSUPPORTOS
ErrorMsg=GetMessage("NOTSUPPORTOS")
Exit Do
End If
If Not isAdmin Then
exitCode=EBP_NOTADMIN
ErrorMsg=GetMessage("NOTADMIN")
Exit Do
End If
If MutexExist(strMutexName)Then
exitCode=EBP_ALREADYEXEC
ErrorMsg=GetMessage("ALREADYEXEC")
Exit Do
End If
Set objHelper=(New CTosWshHelper)(x64Mode,strMutexName)
If Not objHelper.Activated Then
exitCode=EBP_CANNOTEXEC
ErrorMsg=GetMessage("CANNOTEXEC")
Exit Do
End If
If debugMode Then PopupMsg"[Exec]"
With objHelper
MachineID=.MachineInfo("MID")
BiosVersion=.MachineInfo("BiosVersion")
EcType=.MachineInfo("EcType")
EcVersion=.MachineInfo("EcVersion")
End With
platformReader=""
If(x64Mode=True)Then
platformReader=objFS.BuildPath(myFolder,strPlatformReader64Name)
Else
platformReader=objFS.BuildPath(myFolder,strPlatformReader32Name)
End If
platform=objWshShell.Run(platformReader,0,True)
If(platform=PLATFORM_FAILED)Then
exitCode=EBP_NOTSUPPORTPC
ErrorMsg=GetMessage("NOTSUPPORTPC")
Exit Do
ElseIf(platform=PLATFORM_UNKNOWN)Then
If(MachineID="00A1")Then
If debugMode Then
PopupMsg"Pegasus10P"
End If
isPegasus10P=True
EcType=strEcTypePegasus10P
Else
exitCode=EBP_NOTSUPPORTPC
ErrorMsg=GetMessage("NOTSUPPORTPC")
Exit Do
End If
ElseIf(platform=PLATFORM_CRESCENTBAY_OR_OLDER)Then
If debugMode Then
PopupMsg"CrescentBay or older"
End If
isCrescentBayOrOlder=True
ElseIf(platform=PLATFORM_SKYLAKE_OR_LATER_CSM)Then
If debugMode Then
PopupMsg"Skylake or later CSM"
End If
isSkylakeOrLaterCSM=True
ElseIf(platform=PLATFORM_SKYLAKE_OR_LATER_UEFI)Then
If debugMode Then
PopupMsg"Skylake or later UEFI"
End If
isSkylakeOrLaterUEFI=True
Else
exitCode=EBP_NOTSUPPORTPC
ErrorMsg=GetMessage("NOTSUPPORTPC")
Exit Do
End If
If debugMode Then
PopupMsg"MachineID = "&MachineID&vbCrLf&"BiosVersion = """&BiosVersion&""""&vbCrLf&"EcType = "&EcType&vbCrLf&"EcVersion = """&EcVersion&""""
End If
If Len(MachineID)<>4 Then MachineID=""
If Len(EcType)<>4 Then EcType=""
If(Len(EcVersion)<>6)Or(Left(EcVersion,1)<>"V")Then EcVersion=""
If(MachineID="")Or(BiosVersion="")Or(EcType="")Or(EcVersion="")Then
exitCode=EBP_NOTSUPPORTPC
ErrorMsg=GetMessage("NOTSUPPORTPC")
Exit Do
End If
BiosFileName="BIO"&MachineID&"T.COM"
If(isPegasus10P=True)Then
BiosFileName="BIO"&MachineID&"T.BIN"
End If
BiosFile=objFS.BuildPath(myFolder,BiosFileName)
BiosPegasus10PToolName=strBiosPegasus10PToolName
BiosPegasus10PTool=objFS.BuildPath(myFolder,BiosPegasus10PToolName)
BiosFolder=objFS.BuildPath(myFolder,strBiosFolderName)
If objFS.FolderExists(BiosFolder)Then
NewBiosFile=objFS.BuildPath(BiosFolder,BiosFileName)
If objFS.FileExists(NewBiosFile)then
If objFS.FileExists(BiosFile)then objFS.DeleteFile BiosFile,True
If debugMode Then
objFS.CopyFile NewBiosFile,BiosFile,False
Else
objFS.MoveFile NewBiosFile,BiosFile
End If
End If
NewBiosPegasus10PTool=objFS.BuildPath(BiosFolder,BiosPegasus10PToolName)
If objFS.FileExists(NewBiosPegasus10PTool)then
If objFS.FileExists(BiosPegasus10PTool)then objFS.DeleteFile BiosPegasus10PTool,True
If debugMode Then
objFS.CopyFile NewBiosPegasus10PTool,BiosPegasus10PTool,False
Else
objFS.MoveFile NewBiosPegasus10PTool,BiosPegasus10PTool
End If
End If
If ownerMode And Not debugMode Then objFS.DeleteFolder BiosFolder,True
End If
If(BiosFile<>"")And Not objFS.FileExists(BiosFile)Then BiosFile=""
If BiosFile="" Then BiosFileName=""
If(BiosPegasus10PTool<>"")And Not objFS.FileExists(BiosPegasus10PTool)Then BiosPegasus10PTool=""
If BiosPegasus10PTool<>"" Then BiosPegasus10PToolName=""
EcFileName=""
EcFile=""
EcPegasus10PToolName=""
EcPegasus10PTool=""
EcFolder=objFS.BuildPath(myFolder,strEcFolderName)
If objFS.FolderExists(EcFolder)Then
For Each objFile In objFS.GetFolder(EcFolder).Files
With objFile
If(Left(.Name,4)=EcType)And(UCase(objFS.GetExtensionName(.Name))="BIN")Then
EcFileName=.Name
EcFile=objFS.BuildPath(myFolder,EcFileName)
If objFS.FileExists(EcFile)then objFS.DeleteFile EcFile,True
If debugMode Then
objFS.CopyFile .Path,EcFile,False
Else
objFS.MoveFile .Path,EcFile
End If
Exit For
End If
If((Left(.Name,2)=Left(EcType,2))And(isPegasus10P=True)And(UCase(objFS.GetExtensionName(.Name))=UCase(strEcPegasus10PFileExt)))Then
EcFileName=.Name
EcFile=objFS.BuildPath(myFolder,EcFileName)
EcPegasus10PToolName=strEcPegasus10PToolName
EcPegasus10PTool=objFS.BuildPath(myFolder,EcPegasus10PToolName)
If objFS.FileExists(EcFile)then objFS.DeleteFile EcFile,True
If debugMode Then
objFS.CopyFile .Path,EcFile,False
Else
objFS.MoveFile .Path,EcFile
End If
NewEcPegasus10PTool=objFS.BuildPath(EcFolder,EcPegasus10PToolName)
If objFS.FileExists(NewEcPegasus10PTool)then
If objFS.FileExists(EcPegasus10PTool)then objFS.DeleteFile EcPegasus10PTool,True
If debugMode Then
objFS.CopyFile NewEcPegasus10PTool,EcPegasus10PTool,False
Else
objFS.MoveFile NewEcPegasus10PTool,EcPegasus10PTool
End If
End If
Exit For
End If
End With
Next
If ownerMode And Not debugMode Then objFS.DeleteFolder EcFolder,True
End If
If EcFile="" Then
For Each objFile In objFS.GetFolder(myFolder).Files
With objFile
If(Left(.Name,4)=EcType)And(UCase(objFS.GetExtensionName(.Name))="BIN")Then
EcFileName=.Name
EcFile=.Path
Exit For
End If
If((Left(.Name,2)=Left(EcType,2))And(isPegasus10P=True)And(UCase(objFS.GetExtensionName(.Name))=UCase(strEcPegasus10PFileExt)))Then
EcFileName=.Name
EcFile=.Path
Exit For
End If
End With
Next
End If
If(EcFile<>"")And Not objFS.FileExists(EcFile)Then EcFile=""
If EcFile="" Then EcFileName=""
If(EcPegasus10PTool<>"")And Not objFS.FileExists(EcPegasus10PTool)Then EcPegasus10PTool=""
If EcPegasus10PTool<>"" Then EcPegasus10PToolName=""
If debugMode Then PopupMsg"BiosFile = """&BiosFile&""""&vbCrLf&"BiosPegasus10PTool = """&BiosPegasus10PTool&""""&vbCrLf&"EcFile = """&EcFile&""""&vbCrLf&"EcPegasus10PTool = """&EcPegasus10PTool&""""
If(BiosFile="")And(EcFile="")then
exitCode=EBP_NOTCOMPATIBLEPC
ErrorMsg=GetMessage("NOTCOMPATIBLEPC")
Exit Do
End If
BiosFileVer=""
EcFileVer=""
strVerMsg=""
If BiosFile<>"" Then
BiosFileVer=GetBiosFileVer(BiosFile,"BIOS")
If debugMode Then PopupMsg"BiosFileVer = """&BiosFileVer&""""
If BiosFileVer<>"" Then
ver1=StringReplace(Mid(BiosVersion,2)," ","  ")
ver2=StringReplace(Mid(BiosFileVer,2)," ","  ")
strVerMsg=strVerMsg&StringReplace(ReplMessage("BIOSVERS","%1",ver1),"%2",ver2)&vbCrLf
Else
strVerMsg="BIOS"
exitCode=EBP_NOTCOMPATIIMAGE
End If
End If
If EcFile<>"" Then
EcFileVer=GetEcFileVer(EcFile,EcType)
If debugMode Then PopupMsg"EcFileVer = """&EcFileVer&""""
If EcFileVer<>"" Then
ver1=StringReplace(Mid(EcVersion,2)," ","  ")
ver2=StringReplace(Mid(EcFileVer,2)," ","  ")
strVerMsg=strVerMsg&StringReplace(ReplMessage("ECVERS","%1",ver1),"%2",ver2)&vbCrLf
Else
If exitCode<>ERROR_SUCCESS Then
strVerMsg=GetMessage("BIOSANDEC")
Else
strVerMsg="EC"
End If
exitCode=EBP_NOTCOMPATIIMAGE
End If
End If
If exitCode<>ERROR_SUCCESS Then
ErrorMsg=ReplMessage("NOTCOMPATIIMAGE","%s",strVerMsg)
Exit Do
End If
UpdateBIOS=BiosFile<>""
UpdateEC=EcFile<>""
strTarget=GetTargetMessage(UpdateBIOS,UpdateEC)
For i=LBound(specialCase)To UBound(specialCase)
Select Case specialCase(i)(0)
Case 1
If((specialCase(i)(1)="")Or(MachineID=specialCase(i)(1)))And((specialCase(i)(2)="")Or(EcType=specialCase(i)(2)))Then
compStr=specialCase(i)(4)
tempStr=Left(EcVersion,Len(compStr))
Select Case specialCase(i)(3)
Case"="
If tempStr=compStr Then needBattChk=False
Case">"
If tempStr>compStr Then needBattChk=False
Case"<"
If tempStr<compStr Then needBattChk=False
Case">="
If tempStr>=compStr Then needBattChk=False
Case"<="
If tempStr<=compStr Then needBattChk=False
Case"<>"
If tempStr<>compStr Then needBattChk=False
End Select
If debugMode And Not needBattChk Then PopupMsg"Special Case 1: needBattChk = "&needBattChk
End If
End Select
Next
If Not fdMode Then
strUpdMsg=""
strNonMsg=""
strErrMsg=""
RewriteBIOS=False
VerDownBIOS=False
RewriteEC=False
VerDownEC=False
If UpdateBIOS then
flg=CompareFirmwareVersion(BiosFileVer,BiosVersion)
If flg=0 Then
If allowReWriteMode Then
RewriteBIOS=True
Else
UpdateBIOS=False
End If
ElseIf flg<0 Then
If allowVerDownMode Then
VerDownBIOS=True
Else
UpdateBIOS=False
End If
End If
End If
If UpdateEC then
flg=CompareFirmwareVersion(EcFileVer,EcVersion)
If flg=0 Then
If allowReWriteMode Then
RewriteEC=True
Else
UpdateEC=False
End If
ElseIf flg<0 Then
If allowVerDownMode Then
VerDownEC=True
Else
UpdateEC=False
End If
End If
End If
BiosVerTooLow=False
If UpdateBIOS Or UpdateEC Then
If(MachineID=strMachineIdPegasus10P)Then
flg=CompareFirmwareVersion(BiosVersion,strRequiredBiosVersionPegasus10P)
If(flg<0)Then
BiosVerTooLow=True
UpdateBIOS=False
UpdateEc=False
End If
End If
End If
If UpdateBIOS Or UpdateEC Then
strTarget=GetTargetMessage(UpdateBIOS,UpdateEC)
If UpdateBIOS And UpdateEC Then
If RewriteBIOS And RewriteEC Then
strUpdMsg=ReplMessage("FORCEREWRITE","%s",strTarget)
ElseIf VerDownBIOS And VerDownEC Then
strUpdMsg=ReplMessage("FORCEVERDOWN","%s",strTarget)
ElseIf RewriteBIOS Or VerDownBIOS Or RewriteEC Or VerDownEC Then
strUpdMsg=ReplMessage("FORCEUPDATE","%s",strTarget)
Else
strUpdMsg=ReplMessage("NEEDUPDATE","%s",strTarget)
End If
ElseIf RewriteBIOS Or RewriteEC Then
strUpdMsg=ReplMessage("FORCEREWRITE","%s",strTarget)
ElseIf VerDownBIOS Or VerDownEC Then
strUpdMsg=ReplMessage("FORCEVERDOWN","%s",strTarget)
Else
strUpdMsg=ReplMessage("NEEDUPDATE","%s",strTarget)
End If
ElseIf(BiosVerTooLow=True)Then
exitCode=EBP_CANNOTEXEC
strNonMsg=GetMessage("CANNOTEXEC")
Else
exitCode=EBP_UPTODATE
If strTarget=GetMessage("BIOSANDEC")Then
strNonMsg=ReplMessage("NONEEDUPDATE2","%s",strTarget)
Else
strNonMsg=ReplMessage("NONEEDUPDATE","%s",strTarget)
End If
End If
If checkOnlyMode Then
If Not suppressPopup Then
CloseExecutingBar
PopupInfoMsg strUpdMsg&strNonMsg
End If
Exit Do
End If
Set objMsgBox=objHelper.CreatePopup
Do
If Not objHelper.IsPopup(objMsgBox)Then
exitCode=EBP_CANNOTEXEC
ErrorMsg=GetMessage("CANNOTEXEC")
Exit Do
End If
If exitCode=ERROR_SUCCESS Then exitCode=EBP_EXECCANCELED
With objMsgBox
If Not suppressPopup Then
If Not needBattChk Then
.Property("Width")=MsgBoxWidth2
.Property("Height")=MsgBoxHeight2
Else
.Property("Width")=MsgBoxWidth
.Property("Height")=MsgBoxHeight
End If
.Property("BaseColor")=FONT_COLOR_WHITE
.Property("Title")=PopupTitle
.ButtonAdd GetMessage("EXECBUTTON"),False
.ButtonAdd GetMessage("CLOSEBUTTON")
.ButtonDefault=POPUP_BUTTON2
.Open
CloseExecutingBar
End If
IfFirstTime=True
CanExec=False
ReqExec=False
ReqExit=suppressPopup
BDEstate=0
ExitSleep=0
CheckCount=0
Do
If CheckCount>0 Then
CheckCount=CheckCount-1
ElseIf CheckCount=0 Then
CheckCount=10
If UpdateBIOS Or UpdateEC Then
strWarnMsg=""
CurState=True
If LCase(objHelper.MachineInfo("iTxt"))="enabled" Then
strWarnMsg=ReplMessage("ITXTENABLED","%s",strTarget)
CurState=False
ElseIf Not objHelper.PowerCheck(needBattChk)Then
If(isPegasus10P=False)Or((isPegasus10P=True)And(EcVersion<>"V0.00 "))Then
If needBattChk Then
strWarnMsg=GetMessage("NEEDPOWER")
Else
strWarnMsg=GetMessage("NEEDACPOWER")
End If
CurState=False
End If
End If
CurBDEstate=GetBDEProtectionStatus
If IfFirstTime Or(CanExec<>CurState)Or(BDEstate<>CurBDEstate)Then
CanExec=CurState
BDEstate=CurBDEstate
If debugMode And IfFirstTime Then PopupMsg"CanExec = "&CanExec&vbCrLf&"BDEstate = "&BDEstate
If CanExec And autoMode Then ReqExec=True
If Not suppressPopup Then
BtnFocus=POPUP_BUTTON2
.Message=strVerMsg&vbCrLf
If CanExec Then
.MessageLineAdd strUpdMsg,FONT_COLOR_BLUE,FONT_STYLE_BOLD
.MessageLineAdd""
.MessageLineAdd GetMessage("CAUTIONS"),FONT_COLOR_MAROON,FONT_STYLE_BOLD
.MessageLineAdd""
If BDEstate=BDE_PROTECTION_NONE Then
BtnFocus=POPUP_BUTTON1
Else
If BDEstate=BDE_PROTECTION_OFF Then
.MessageLineAdd ReplMessage("BDESUSPENDED","%s",strTarget),FONT_COLOR_MAROON,FONT_STYLE_BOLD
BtnFocus=POPUP_BUTTON1
ElseIf BDEstate=BDE_PROTECTION_ON Then
.MessageLineAdd ReplMessage("BDEENABLED","%s",strTarget),FONT_COLOR_RED,FONT_STYLE_BOLD
Else
.MessageLineAdd ReplMessage("BDEUNKNOWN","%s",strTarget),FONT_COLOR_MAROON,FONT_STYLE_BOLD
End If
.MessageLineAdd""
End If
.MessageLineAdd ReplMessage("NOTICE","%s",strTarget)
If Not needBattChk Then
.MessageLineAdd""
.MessageLineAdd GetMessage("NEEDBATTCHARGE"),FONT_COLOR_MAROON,FONT_STYLE_BOLD
End If
Else
.MessageLineAdd strWarnMsg,FONT_COLOR_RED,FONT_STYLE_BOLD
End If
.Property("MessageScroll")="Top"
.ButtonEnable(POPUP_BUTTON1)=CanExec
.ButtonSetFocus BtnFocus
End If
End If
ElseIf IfFirstTime And Not suppressPopup Then
.Message=strVerMsg&vbCrLf
If strNonMsg<>"" Then
If(BiosVerTooLow=False)Then
.MessageLineAdd strNonMsg,FONT_COLOR_GREEN,FONT_STYLE_BOLD
Else
.MessageLineAdd strNonMsg,FONT_COLOR_ORANGE,FONT_STYLE_BOLD
End If
End If
If autoMode And Not WarnTypeS Then
ExitSleep=3
ReqExit=True
End If
End If
IfFirstTime=False
End If
If Not suppressPopup Then
Button=.ButtonClick
With objHelper
If .OnButtonClick(Button,POPUP_BUTTON2)Then
ReqExit=True
ElseIf CanExec Then
If .OnButtonClick(Button,POPUP_BUTTON1)Then ReqExec=True
End If
End With
End If
If ReqExec Then
CheckCount=-1
CanExec=False
ReqExec=False
If Not suppressPopup Then
.ButtonEnable(POPUP_BUTTON1)=False
.ButtonEnable(POPUP_BUTTON2)=False
.Message=GetMessage("UPDATEPREPARING")
End If
exeParam=""
If(isPegasus10P=True)Then
isESPErrorScheme=True
If x64Mode Then
JoinParam exeParam,strExeESP64Name,False
JoinParam exeParam,"-p",False
Else
JoinParam exeParam,strExeESP32Name,False
JoinParam exeParam,"-p",False
End If
ElseIf(isSkylakeOrLaterCSM=True)Then
If x64Mode Then
JoinParam exeParam,strExe64Name3,False
Else
JoinParam exeParam,strExe32Name3,False
End If
ElseIf(isSkylakeOrLaterUEFI=True)Then
isESPErrorScheme=True
If x64Mode Then
JoinParam exeParam,strExeESP64Name,False
Else
JoinParam exeParam,strExeESP32Name,False
End If
Else
If x64Mode Then
JoinParam exeParam,strExe64Name,False
Else
JoinParam exeParam,strExe32Name,False
End If
End If
If UpdateBIOS Then JoinParam exeParam,BiosFileName,False
If UpdateEC Then JoinParam exeParam,EcFileName,False
If SvPass<>"" Then JoinParam exeParam,"/p="&SvPass,False
ownerHWND=.Property("Handle")
If suppressPopup Then
JoinParam exeParam,"/s",False
ElseIf Len(ownerHWND)=8 Then
If(isCrescentBayOrOlder=True)Then
JoinParam exeParam,"/w="&ownerHWND,False
End IF
End If
If debugMode Then PopupMsg"exeParam = ["&exeParam&"]"
With objHelper.CreateProcess
.CurrentDirectory=myFolder
.Exec exeParam,debugMode,False
Do While .Status=0
If debugMode Then PopupMsg".Status = "&.Status
objMsgBox.Sleep 100
Loop
errCode=.ExitCode
End With
If debugMode Then PopupMsg"errCode = "&errCode
If errCode=0 Then
If objFS.FileExists(objFS.BuildPath(myFolder,strPostExeName))Then
exeParam=""
JoinParam exeParam,strPostExeName,False
If debugMode Then PopupMsg"exeParam = ["&exeParam&"]"
With objHelper.CreateProcess
.CurrentDirectory=myFolder
.Exec exeParam,True,False
Do While .Status=0
objMsgBox.Sleep 100
Loop
errCode=.ExitCode
End With
If debugMode Then PopupMsg"errCode = "&errCode
End If
exitCode=ERROR_SUCCESS
If Not noRebootMode Then rebootMode=True
If Not suppressPopup Then
.Message=""
If noRebootMode Then
.MessageLineAdd GetMessage("NEEDREBOOT"),FONT_COLOR_BLUE,FONT_STYLE_BOLD
If autoMode Then
ExitSleep=5
ReqExit=True
Else
.ButtonEnable(POPUP_BUTTON2)=True
End If
Else
.MessageLineAdd GetMessage("REBOOTING")
ExitSleep=3
ReqExit=True
End If
End If
Else
If isESPErrorScheme Then
If errCode>&HFFF Then
exitCode=EBP_ESPERROR_BASE+&HFFF
Else
exitCode=EBP_ESPERROR_BASE+errCode
End If
strErrMsg=ReplMessage("SOMEERROR","%d","0x"&FormatToHex(Hex(errCode),3))
Else
Select Case errCode
Case 1
exitCode=EBP_BADDEVDRIVER
strErrMsg=GetMessage("BADDEVDRIVER")
Case 2
exitCode=EBP_NOPERMISSION
strErrMsg=GetMessage("NOPERMISSION")
Case 3
exitCode=EBP_CANNOTREADIMAGE
strErrMsg=ReplMessage("CANNOTREADIMAGE","%s","BIOS")
Case 4
exitCode=EBP_CORRUPTEDIMAGE
strErrMsg=ReplMessage("CORRUPTEDIMAGE","%s","BIOS")
Case 5
exitCode=EBP_CANNOTREADIMAGE
strErrMsg=ReplMessage("CANNOTREADIMAGE","%s","EC")
Case 6
exitCode=EBP_CORRUPTEDIMAGE
strErrMsg=ReplMessage("CORRUPTEDIMAGE","%s","EC")
Case 7
exitCode=EBP_INCOMPATIIMAGE
strErrMsg=GetMessage("INCOMPATIIMAGE")
Case 8
exitCode=EBP_BADENCPASFORMAT
strErrMsg=GetMessage("BADENCPASFORMAT")
Case Else
If errCode>&HFF Then
exitCode=EBP_SOMEERROR_BASE+&HFF
Else
exitCode=EBP_SOMEERROR_BASE+errCode
End If
strErrMsg=ReplMessage("SOMEERROR","%d","0x"&FormatToHex(Hex(errCode),2))
End Select
End If
If Not suppressPopup Then
.Message=""
.MessageLineAdd GetMessage("UPDATEFAILED"),FONT_COLOR_RED,FONT_STYLE_BOLD
.ButtonEnable(POPUP_BUTTON2)=True
End If
End If
End If
If Not suppressPopup Then
If strErrMsg<>"" Then .MessageLineAdd strErrMsg,FONT_COLOR_RED,FONT_STYLE_BOLD
If ReqExit Then
.ButtonEnable(POPUP_BUTTON1)=False
.ButtonEnable(POPUP_BUTTON2)=False
If ExitSleep>0 Then .Sleep 1000*ExitSleep
Else
.Sleep 50
End If
End If
strErrMsg=""
Loop Until ReqExit
If Not suppressPopup Then .Close
End With
Loop Until True
End If
If fdMode Then
Set objMsgBox=objHelper.CreatePopup
Do
If Not objHelper.IsPopup(objMsgBox)Then
exitCode=EBP_CANNOTEXEC
ErrorMsg=GetMessage("CANNOTEXEC")
Exit Do
End If
With objMsgBox
.Property("Width")=FdBoxWidth
.Property("Height")=FdBoxHeight
.Property("BaseColor")=FONT_COLOR_WHITE
.Property("Title")=PopupTitle
.ButtonAdd GetMessage("CLOSEBUTTON"),False
.Open
CloseExecutingBar
.Message=ReplMessage("EXTRACTIMAGE","%s",strTarget)
target=targetFolder
ownerHWND=0
Set objAppShell=WScript.CreateObject("Shell.Application")
Set objFolder=Nothing
Do While True
If target="" Then
ownerHWND=CLng("&H"&.Property("Handle"))
Set objFolder=objAppShell.BrowseForFolder(ownerHWND,ReplMessage("SELECTFOLDER","%s",strTarget),&H53)
If objFolder Is Nothing Then
target=""
ElseIf Not objFolder.Self.IsFileSystem Then
target=strUndefinedFolder
If debugMode Then PopupMsg"Undefined Folder = """&objFolder.Self.Path&""""
Else
target=objFolder.Self.Path
End If
Set objFolder=Nothing
End If
If debugMode Then PopupMsg"target = """&target&""""
If target=strUndefinedFolder Then
.MsgBox GetMessage("ILLEGALFOLDER"),vbOkOnly Or vbCritical,PopupTitle
ElseIf target<>"" Then
If Not objFS.FolderExists(target)Then
.MsgBox ReplMessage("UNKNOWNFOLDER","%f",target),vbOkOnly Or vbCritical,PopupTitle
Else
If .MsgBox(ReplMessage(ReplMessage("DOEXTRACT","%s",strTarget),"%f",target),vbYesNo Or vbQuestion,PopupTitle)=vbYes Then Exit Do
End If
Else
exitCode=EBP_EXTRACTCANCELED
Exit Do
End If
target=""
Loop
Set objAppShell=Nothing
If exitCode=ERROR_SUCCESS Then
ExitSleep=0
strExtracted=""
With objFS
ChgBiosFile=.BuildPath(myFolder,strChgBiosName)
If .FileExists(ChgBiosFile)Then
destFile=.BuildPath(target,strChgBiosName)
.CopyFile ChgBiosFile,destFile,True
If .FileExists(destFile)Then
strExtracted=strExtracted&vbCrLf&"  - "&strChgBiosName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
ChgBiosEFIFile=.BuildPath(myFolder,strChgBiosEFIName)
If .FileExists(ChgBiosEFIFile)Then
destEFIFile=.BuildPath(target,strChgBiosEFIName)
.CopyFile ChgBiosEFIFile,destEFIFile,True
If .FileExists(destEFIFile)Then
strExtracted=strExtracted&vbCrLf&"  - "&strChgBiosEFIName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
ChgBiosPegasus10PFile=.BuildPath(myFolder,strChgBiosPegasus10PName)
If .FileExists(ChgBiosPegasus10PFile)Then
destPegasus10PFile=.BuildPath(target,strChgBiosPegasus10PName)
.CopyFile ChgBiosPegasus10PFile,destPegasus10PFile,True
If .FileExists(destPegasus10PFile)Then
strExtracted=strExtracted&vbCrLf&"  - "&strChgBiosPegasus10PName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
ChgBiosPegasus10PTool=.BuildPath(myFolder,strChgBiosPegasus10PToolName)
If .FileExists(ChgBiosPegasus10PTool)Then
destPegasus10PTool=.BuildPath(target,strChgBiosPegasus10PToolName)
.CopyFile ChgBiosPegasus10PTool,destPegasus10PTool,True
If .FileExists(destPegasus10PTool)Then
strExtracted=strExtracted&vbCrlf&"  - "&strChgBiosPegasus10PToolName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
If(BiosFile<>"")And .FileExists(BiosFile)Then
destFile=.BuildPath(target,BiosFileName)
.CopyFile BiosFile,destFile,True
If .FileExists(destFile)Then
strExtracted=strExtracted&vbCrLf&"  - "&BiosFileName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
If(BiosPegasus10PTool<>"")And .FileExists(BiosPegasus10PTool)Then
destPegasus10PTool=.BuildPath(target,BiosPegasus10PToolName)
.CopyFile BiosPegasus10PTool,destPegasus10PTool,True
If .FileExists(destPegasus10PTool)Then
strExtracted=strExtracted&vbCrLf&"  - "&BiosPegasus10PToolName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
If(EcFile<>"")And .FileExists(EcFile)Then
destFile=.BuildPath(target,EcFileName)
.CopyFile EcFile,destFile,True
If .FileExists(destFile)Then
strExtracted=strExtracted&vbCrLf&"  - "&EcFileName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
If(EcPegasus10PTool<>"")And .FileExists(EcPegasus10PTool)Then
destPegasus10PTool=.BuildPath(target,EcPegasus10PToolName)
.CopyFile EcPegasus10PTool,destPegasus10PTool,True
If .FileExists(destPegasus10PTool)Then
strExtracted=strExtracted&vbCrLf&"  - "&EcPegasus10PToolName
Else
exitCode=EBP_CANNOTEXTRACT
End If
End If
End With
.Message=""
If exitCode=ERROR_SUCCESS Then
.MessageLineAdd ReplMessage("EXTRACTED","%f",target)
.MessageLineAdd strExtracted
.Property("MessageScroll")="Top"
Else
strErrMsg=ReplMessage(ReplMessage("CANNOTEXTRACT","%s",strTarget),"%f",target)
.MessageLineAdd strErrMsg,FONT_COLOR_RED,FONT_STYLE_BOLD
End If
Else
ExitSleep=3000
.MessageLineAdd vbCrLf&GetMessage("EXTRACTCANCELED")
End If
.ButtonEnable(POPUP_BUTTON1)=True
.ButtonSetFocus POPUP_BUTTON1
.Sleep ExitSleep,True
.Close
End With
Loop Until True
End If
Loop Until True
Set objMsgBox=Nothing
Set objHelper=Nothing
If(ErrorMsg<>"")And Not suppressPopup Then PopupErrMsg ErrorMsg
If debugMode And cleanMode Then
If MsgBox("Debug: Clean now ?",vbYesNo Or vbQuestion Or vbDefaultButton2 Or vbApplicationModal,PopupTitle)=vbNo Then cleanMode=False
End If
If cleanMode Then
If debugMode Then PopupMsg"[Clean]"
If Not Cleaning(myFolder,True)Then
exitCode=EBP_CANNOTDELFOLDER
If Not suppressPopup Then PopupWarnMsg ReplMessage("CANNOTDELFOLDER","%f",myFolder)
End If
If objFS.FileExists(myPath)Then objFS.DeleteFile myPath,True
End If
If(exitCode<>ERROR_SUCCESS)And(exitCode<>EBP_CANNOTDELFOLDER)Then rebootMode=False
If exitCode<>ERROR_SUCCESS Then exitCode=ERROR_EBP_BASE+exitCode
If debugMode And rebootMode Then
If MsgBox("Debug: Reboot now ?",vbYesNo Or vbQuestion Or vbDefaultButton2 Or vbApplicationModal,PopupTitle)=vbNo Then rebootMode=False
End If
Set objMessage=Nothing
Set objWshShell=Nothing
Set objFS=Nothing
If rebootMode Then Reboot
WScript.Quit exitCode
Function FormatToHex(ByVal str,ByVal n)
If Len(str)>=n Then
FormatToHex=str
Else
FormatToHex=String(n-Len(str),"0")&str
End If
End Function
Function StringCheck(ByVal str,ByVal patrn,ByVal ignCase)
Dim reg
Set reg=New RegExp
With reg
.Pattern=patrn
.Global=False
.IgnoreCase=ignCase
.MultiLine=True
StringCheck=.Test(str)
End With
Set reg=Nothing
End Function
Function StringReplace(ByVal str,ByVal patrn,ByVal replStr)
Dim reg
Set reg=New RegExp
With reg
.Pattern=patrn
.Global=True
.IgnoreCase=False
.MultiLine=True
StringReplace=.Replace(str,replStr)
End With
Set reg=Nothing
End Function
Sub TranslateMessages
Dim file,str,key,msg
With objFS
file=.BuildPath(myFolder,.GetBaseName(myName)&"."&LangID)
If Not .FileExists(file)Then file=.BuildPath(myFolder,.GetBaseName(myName)&"."&strDefLangId)
If .FileExists(file)Then
With .OpenTextFile(file)
Do Until .AtEndOfStream
str=.ReadLine
If Left(LTrim(StringReplace(str,"\t","")),1)=";" Then str=""
If InStr(str,"=")>1 Then
key=Trim(StringReplace(Split(str,"=",2)(0),"\t",""))
msg=Split(str,"=",2)(1)
msg=StringReplace(msg,"\\n",vbCrLf)
msg=StringReplace(msg,"\\t",vbTab)
With objMessage
If .Exists(key)Then
.Key(key)=msg
Else
.Add key,msg
End If
End With
End If
Loop
End With
End If
End With
End Sub
Function GetMessage(ByVal Key)
GetMessage=Key
With objMessage
If .Exists(Key)Then GetMessage=.Item(Key)
End With
End Function
Function ReplMessage(ByVal Key,ByVal patrn,ByVal replStr)
ReplMessage=StringReplace(GetMessage(Key),patrn,replStr)
End Function
Function GetTargetMessage(ByVal BIOS,ByVal EC)
If BIOS And EC Then
GetTargetMessage=GetMessage("BIOSANDEC")
ElseIf BIOS Then
GetTargetMessage="BIOS"
Else
GetTargetMessage="EC"
End If
End Function
Sub CloseExecutingBar
Dim CloseFlgFile
If ExecutingBarClosed Then Exit Sub
CloseFlgFile=objFS.BuildPath(myFolder,strCloseFlgName)
If Not objFS.FileExists(CloseFlgFile)Then
If debugMode Then objWshShell.Popup"CloseExecutingBar",vbOkOnly Or vbApplicationModal
objFS.CreateTextFile(CloseFlgFile,True).Close
End If
ExecutingBarClosed=True
End Sub
Sub PopupMsg(ByVal msg)
objWshShell.Popup msg,0,PopupTitle,vbOkOnly Or vbApplicationModal
End Sub
Sub PopupTimeMsg(ByVal msg,timeout)
objWshShell.Popup msg,timeout,PopupTitle,vbOkOnly Or vbApplicationModal
End Sub
Sub PopupInfoMsg(ByVal msg)
objWshShell.Popup msg,0,PopupTitle,vbOkOnly Or vbInformation Or vbApplicationModal
End Sub
Sub PopupWarnMsg(ByVal msg)
objWshShell.Popup msg,0,PopupTitle,vbOkOnly Or vbExclamation Or vbApplicationModal
End Sub
Sub PopupErrMsg(ByVal msg)
Dim reqClose
reqClose=(ownerMode And(exitCode<>ERROR_SUCCESS))
If(exitCode<>EBP_CANNOTDELFOLDER)And cleanMode And Not debugMode Then Cleaning myFolder,Not reqClose
If reqClose Then CloseExecutingBar
objWshShell.Popup msg,0,PopupTitle,vbOkOnly Or vbCritical Or vbApplicationModal
End Sub
Sub JoinParam(param,ByVal str,ByVal unshift)
If str<>"" Then
If InStr(str," ")>=1 Then str=""""&str&""""
If unshift Then
If param<>"" Then param=" "&param
param=str&param
Else
If param<>"" Then param=param&" "
param=param&str
End If
End If
End Sub
Function GetIniKeyVal(ByVal iniFile,ByVal section,ByVal key)
Dim flg,str,aStr
GetIniKeyVal=""
If objFS.FileExists(iniFile)then
With objFS.OpenTextFile(iniFile)
flg=(section="")
section=LCase("["&section&"]")
key=LCase(key)
Do Until .AtEndOfStream
str=LTrim(.ReadLine)
If Not flg Then
If LCase(Trim(str))=section Then flg=True
Else
If Left(str,1)="[" Then Exit Do
aStr=Split(str,"=",2)
If(UBound(aStr)=1)And(LCase(Trim(aStr(0)))=key)Then
GetIniKeyVal=Trim(aStr(1))
If(Len(GetIniKeyVal)>1)And(Left(GetIniKeyVal,1)="""")And(Right(GetIniKeyVal,1)="""")Then
GetIniKeyVal=Mid(GetIniKeyVal,2,Len(GetIniKeyVal)-2)
End If
Exit Do
End If
End If
Loop
End With
End If
End Function
Const OS_Unknown=0
Const OS_Win95=1
Const OS_Win98=2
Const OS_WinNT4=3
Const OS_Win2000=4
Const OS_WinXP=5
Const OS_WinSvr2003=6
Const OS_WinVista=7
Const OS_Win7=8
Const OS_Win8=9
Const OS_Win8_1=10
Const OS_Win10=11
Const OS_WinNTbase=99
Dim OSTypeCache:OSTypeCache=OS_Unknown
Function GetOSType
Dim objWMI,objSys
GetOSType=OSTypeCache
If OSTypeCache>OS_Unknown Then Exit Function
On Error Resume Next
Set objWMI=GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
For Each objSys in objWMI.ExecQuery("Select * from Win32_OperatingSystem")
With objSys
Select Case .OStype
Case 16
OSTypeCache=OS_Win95
Case 17
OSTypeCache=OS_Win98
Case 18
Select Case Left(.Version,3)
Case"4.0"
OSTypeCache=OS_WinNT4
Case"5.0"
OSTypeCache=OS_Win2000
Case"5.1"
OSTypeCache=OS_WinXP
Case"5.2"
OSTypeCache=OS_WinSvr2003
Case"6.0"
OSTypeCache=OS_WinVista
Case"6.1"
OSTypeCache=OS_Win7
Case"6.2"
OSTypeCache=OS_Win8
Case"6.3"
OSTypeCache=OS_Win8_1
Case"10."
OSTypeCache=OS_Win10
Case Else
OSTypeCache=OS_WinNTbase
End Select
Case Else
OSTypeCache=OS_Unknown
End Select
End With
Next
On Error Goto 0
Set objWMI=Nothing
GetOSType=OSTypeCache
End Function
Function GetByte(buf,pos)
GetByte=AscB(MidB(buf,pos,1))
End Function
Function BinToStr(ByVal buf,ByVal size)
Dim i,str
str=""
If size>0 Then
For i=1 to size
str=str&Chr(GetByte(buf,i))
Next
End If
BinToStr=str
End Function
Function GetBiosFileVer(ByVal target,ByVal signature)
Const adTypeBinary=1
Dim val,ver
ver=""
If objFS.FileExists(target)Then
With WScript.CreateObject("ADODB.Stream")
.Open
.Type=adTypeBinary
.LoadFromFile target
.Position=&H00
val=.Read(3)
If(GetByte(val,1)=0)And(GetByte(val,2)=0)And(GetByte(val,3)<=2)Then
.Position=&H03
If BinToStr(.Read(4),4)=signature Then
.Position=&H0B
val=BinToStr(.Read(6),6)
If(Len(val)=6)And(Left(val,1)="v")Then ver=val
End If
End If
.Close
End With
End If
If objFS.FileExists(target)And(LCase(Right(target,Len(strBiosPegasus10PFileExt)))=LCase(strBiosPegasus10PFileExt))Then
With WScript.CreateObject("ADODB.Stream")
.Open
.Type=adTypeBinary
.LoadFromFile target
.Position=&H417200
val=.Read(6)
If BinToStr(val,6)="$BVDT$" Then
.Position=&H41720D
val=.Read(9)
If BinToStr(val,9)="$Version " Then
.Position=&H417216
val=BinToStr(.Read(5),5)
if(Len(val)=5)Then
ver="v"+val
End If
End If
End If
.Close
End With
End If
GetBiosFileVer=ver
End Function
Function GetEcFileVer(ByVal target,ByVal signature)
Const adTypeBinary=1
Dim val,ver
ver=GetBiosFileVer(target,signature)
If ver="" Then
If objFS.FileExists(target)Then
With WScript.CreateObject("ADODB.Stream")
.Open
.Type=adTypeBinary
.LoadFromFile target
.Position=&H80
If BinToStr(.Read(8),8)="progdown" Then
.Position=&H88
If GetByte(.Read(1),1)=&H02 Then
.Position=&H8A
val=BinToStr(.Read(6),6)
If(Len(val)=6)And(Left(val,1)="V")Then ver=val
End If
End If
.Close
End With
End If
If objFS.FileExists(target)And(LCase(Right(target,Len(strEcPegasus10PFileExt)))=LCase(strEcPegasus10PFileExt))Then
With WScript.CreateObject("ADODB.Stream")
.Open
.Type=adTypeBinary
.LoadFromFile target
.Position=&H0
If BinToStr(.Read(5),5)="@a000" Then
.Position=&H7
ver=BinToStr(.Read(9),9)
if(Mid(ver,1,2)="00")And(Mid(ver,7,1)="0")And(Mid(ver,3,1)=" ")And(Mid(ver,6,1)=" ")And(Mid(ver,9,1)=" ")Then
ver="V"&Mid(ver,8,1)&"."&Mid(ver,4,2)
End If
End If
.Close
End With
End If
End If
GetEcFileVer=ver
End Function
Const BDE_PROTECTION_NONE=-1
Const BDE_PROTECTION_OFF=0
Const BDE_PROTECTION_ON=1
Const BDE_PROTECTION_UNKNOWN=2
Const BDE_FULLY_ENCRYPTED=1
Function GetBDEProtectionStatus
Dim objWMIService,volumes,volume,sysDrive,res,val,val2
GetBDEProtectionStatus=BDE_PROTECTION_NONE
On Error Resume Next
If OSType>=OS_WinVista Then
Set objWMIService=GetObject("winmgmts:\\.\root\CIMV2\Security\MicrosoftVolumeEncryption")
If Err.Number=0 Then
Set volumes=objWMIService.InstancesOf("Win32_EncryptableVolume")
If Err.Number=0 Then
GetBDEProtectionStatus=BDE_PROTECTION_UNKNOWN
sysDrive=objWshShell.ExpandEnvironmentStrings("%SystemDrive%")
If(Len(sysDrive)<>2)Or(Right(sysDrive,1)<>":")Then sysDrive="C:"
For Each volume In volumes
If volume.DriveLetter=sysDrive Then
res=volume.GetProtectionStatus(val)
If res=0 Then
If val=BDE_PROTECTION_OFF Then
res=volume.GetConversionStatus(val,val2)
If res=0 Then
If val=BDE_FULLY_ENCRYPTED Then
GetBDEProtectionStatus=BDE_PROTECTION_OFF
Else
GetBDEProtectionStatus=BDE_PROTECTION_NONE
End If
Else
GetBDEProtectionStatus=BDE_PROTECTION_UNKNOWN
End If
Else
GetBDEProtectionStatus=val
End If
Else
GetBDEProtectionStatus=BDE_PROTECTION_UNKNOWN
End If
Exit For
End If
Next
End If
Set volumes=Nothing
End If
Set objWMIService=Nothing
End If
On Error Goto 0
End Function
Function Cleaning(ByVal cleanFolder,ByVal delMe)
Dim errFlg,delMeFlg,UCcurFolder,UCcleanFolder,retryCount,loopCount
Dim objFolder,objSubFolder,objFile,target
errFlg=False
delMeFlg=delMe
On Error Resume Next
With objFS
If .FolderExists(cleanFolder)Then
If debugMode Then PopupMsg"Cleaning: cleanFolder = "&cleanFolder
errFlg=.GetFolder(cleanFolder).IsRootFolder
If Not errFlg Then
errFlg=(Not StringCheck(cleanFolder,"\\_cab.+\.tmp$",True))And(Not StringCheck(cleanFolder,"\\"&strToshibaFolder&"\\"&strPackageFolder&"$",True))
End If
If Not errFlg Then
Set objMsgBox=Nothing
Set objHelper=Nothing
UCcurFolder=UCase(curFolder)
UCcleanFolder=UCase(cleanFolder)
While Instr(UCcurFolder,UCcleanFolder)=1
curFolder=.GetParentFolderName(curFolder)
UCcurFolder=UCase(curFolder)
Wend
If Not .FolderExists(curFolder)Then curFolder=.GetSpecialFolder(2).Path
If .FolderExists(curFolder)Then objWshShell.CurrentDirectory=curFolder
If debugMode Then PopupMsg"Cleaning: Current folder = "&objWshShell.CurrentDirectory
retryCount=5
Do
Set objFolder=.GetFolder(cleanFolder)
For Each objSubFolder In objFolder.SubFolders
target=objSubFolder.Path
objSubFolder.Delete True
If .FolderExists(target)Then errFlg=True
Next
For Each objFile In objFolder.Files
target=objFile.Path
If objFile.Name=strCloseFlgName Then
loopCount=5
Do
WScript.Sleep(50)
If Not .FileExists(target)Then Exit Do
loopCount=loopCount-1
Loop While loopCount>0
End If
If .FileExists(target)Then
objFile.Delete True
If .FileExists(target)Then errFlg=True
End If
Next
If delMeFlg Then
.DeleteFolder cleanFolder,True
If .FolderExists(cleanFolder)Then errFlg=True
End If
If errFlg Then
retryCount=retryCount-1
If debugMode Then PopupMsg"Cleaning: retryCount = "&retryCount
If retryCount=0 Then Exit Do
errFlg=False
WScript.Sleep(1000)
Else
If debugMode Then PopupMsg"Cleaning: Succeeded."
Exit Do
End If
Loop While True
End If
End If
End With
On Error Goto 0
Cleaning=(Not errFlg)
If debugMode Then PopupMsg"Cleaning: res = "&Cleaning
End Function
Sub Reboot
For Each objSys In GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}").InstancesOf("Win32_OperatingSystem")
objSys.Win32Shutdown 2
Next
End Sub
Const CTWH_Server32Name="TosWshHelper.exe"
Const CTWH_Server64Name="TosWshHelper64.exe"
Const CTWH_ClassName="TosWshHelper"
Const CTWH_IMutex="Mutex"
Const CTWH_IPopup="Popup"
Const CTWH_IMachine="Machine"
Const CTWH_IProcess="Process"
Const POPUP_BUTTON1=1
Const POPUP_BUTTON2=2
Const POPUP_BUTTON3=3
Const FONT_STYLE_NONE=&H00
Const FONT_STYLE_BOLD=&H01
Const FONT_STYLE_ITALIC=&H02
Const FONT_STYLE_UNDERLINE=&H04
Const FONT_STYLE_STRIKEOUT=&H08
Const FONT_STYLE_BLINK=&H80
Class CTosWshHelper
Private Server,bComFound,bComActivated
Private objMutex,objMachine
Private Sub Class_Initialize
Server=""
bComFound=False
bComActivated=False
End Sub
Public Default Function Init(ByVal x64,ByVal MutexName)
Dim ClsidKey,Clsid,BaseKey,TypeLibKey,VersionKey,LibVer,LibPath,LibID,LibKey
Dim ServerPath,ServerVer,exeParam,FolderBase,FolderName
Set Init=Me
If bComActivated Or bComFound Or(Server<>"")Then Exit Function
With objFS
If x64 Then ServerPath=CTWH_Server64Name Else ServerPath=CTWH_Server32Name
ServerPath=.BuildPath(.GetParentFolderName(WScript.ScriptFullName),ServerPath)
ClsidKey="SOFTWARE\Classes\"&CTWH_ClassName&"."&CTWH_IMutex&"\Clsid"
If .FileExists(ServerPath)And ExistsRegKey(RegHKLM,ClsidKey)Then
ServerVer=.GetFileVersion(ServerPath)
Clsid=GetRegStringValue(RegHKLM,ClsidKey,"")
BaseKey="SOFTWARE\Wow6432Node\Classes"
If Not ExistsRegKey(RegHKLM,BaseKey)Then BaseKey="SOFTWARE\Classes"
TypeLibKey=BaseKey&"\CLSID\"&Clsid&"\TypeLib"
VersionKey=BaseKey&"\CLSID\"&Clsid&"\Version"
LibVer=""
LibPath=""
If ExistsRegKey(RegHKLM,TypeLibKey)And ExistsRegKey(RegHKLM,VersionKey)Then
LibID=GetRegStringValue(RegHKLM,TypeLibKey,"")
LibVer=GetRegStringValue(RegHKLM,VersionKey,"")
LibKey=BaseKey&"\TypeLib\"&LibID&"\"&LibVer&"\0\win32"
If ExistsRegKey(RegHKLM,LibKey)Then LibPath=GetRegStringValue(RegHKLM,LibKey,"")
End If
If(InStr(ServerVer,".")>1)And(InStr(LibVer,".")>1)And .FileExists(LibPath)Then
If CompareVersion(ServerVer,LibVer)>0 Then
If RegServer(LibPath,False)Then
FolderBase=.GetParentFolderName(LibPath)
FolderName=LCase(.GetFileName(FolderBase))
If(Left(FolderName,4)="_cab")And(Right(FolderName,4)=".tmp")Then
On Error Resume Next
.DeleteFolder FolderBase,True
On Error Goto 0
End If
End If
End If
End If
End If
End With
With WScript
Server=""
bComFound=False
bComActivated=False
On Error Resume Next
Set objMutex=.CreateObject(CTWH_ClassName&"."&CTWH_IMutex)
If Err.Number=0 Then bComFound=True
On Error Goto 0
If Not bComFound Then
If RegServer(ServerPath,True)Then Server=ServerPath
If Server<>"" Then
On Error Resume Next
Set objMutex=.CreateObject(CTWH_ClassName&"."&CTWH_IMutex)
If Err.Number=0 Then bComFound=True
On Error Goto 0
End If
End If
If bComFound Then
On Error Resume Next
bComActivated=objMutex.Create(MutexName)
On Error Goto 0
If bComActivated Then
On Error Resume Next
Set objMachine=.CreateObject(CTWH_ClassName&"."&CTWH_IMachine)
If Err Then Set objMachine=Nothing
On Error Goto 0
End If
End If
End With
End Function
Private Sub Class_Terminate
Dim n
On Error Resume Next
If bComFound Then
Set objMachine=Nothing
objMutex.Close
Set objMutex=Nothing
End If
RegServer Server,False
On Error Goto 0
End Sub
Public Property Get Activated
Activated=bComActivated
End Property
Public Function IsPopup(objPopup)
IsPopup=(TypeName(objPopup)=CTWH_IPopup)
End Function
Public Function CreatePopup
Set CreatePopup=Nothing
On Error Resume Next
Set CreatePopup=WScript.CreateObject(CTWH_ClassName&"."&CTWH_IPopup)
On Error Goto 0
If Not IsPopup(CreatePopup)Then Set CreatePopup=Nothing
End Function
Public Function OnButtonClick(ByVal Click,ByVal Num)
If(Click<>0)And(Num>0)Then
OnButtonClick=((Click And(2^(Num-1)))<>0)
Else
OnButtonClick=False
End If
End Function
Public Function IsMachine(objMachine)
IsMachine=(TypeName(objMachine)=CTWH_IMachine)
End Function
Public Function MachineInfo(ByVal Name)
MachineInfo=""
If bComFound And IsMachine(objMachine)Then
On Error Resume Next
With objMachine
MachineInfo=.Info(Name)
End With
On Error Goto 0
End If
End Function
Public Function UpdatableFW(ByVal BIOSver,ByVal ECver)
UpdatableFW=True
If bComFound And IsMachine(objMachine)Then
On Error Resume Next
If objMachine.Func("CompFW",ECver,BIOSver)="False" Then UpdatableFW=False
On Error Goto 0
End If
End Function
Public Function PowerCheck(ByVal withBattChk)
Dim res,ACLineStatus,BatteryLifePercent
ACLineStatus=MachineInfo("ACLineStatus")
res=(ACLineStatus="Online")
If withBattChk Then
BatteryLifePercent=MachineInfo("BatteryLifePercent")
If BatteryLifePercent="Unknown" Then BatteryLifePercent=0
res=res And(StrToInt(BatteryLifePercent)>=MinimumBatteryLifePercent)
End If
PowerCheck=res
End Function
Public Function IsProcess(objProcess)
IsProcess=(TypeName(objProcess)=CTWH_IProcess)
End Function
Public Function CreateProcess
Set CreateProcess=Nothing
On Error Resume Next
Set CreateProcess=WScript.CreateObject(CTWH_ClassName&"."&CTWH_IProcess)
On Error Goto 0
If Not IsProcess(CreateProcess)Then Set CreateProcess=Nothing
End Function
End Class
Function MutexExist(ByVal MutexName)
Dim objMutex
MutexExist=False
On Error Resume Next
Set objMutex=WScript.CreateObject(CTWH_ClassName&"."&CTWH_IMutex)
If Err.Number=0 Then
If objMutex.Create(MutexName)Then
objMutex.Close
Else
MutexExist=True
End If
Set objMutex=Nothing
End If
On Error Goto 0
End Function
Function RegServer(ByVal Server,ByVal fRegister)
Dim Regsvr,exeParam
RegServer=False
If Server<>"" then
With objFS
exeParam=""
If LCase(.GetExtensionName(Server))="exe" Then
If .FileExists(Server)then
JoinParam exeParam,Server,False
If fRegister Then
JoinParam exeParam,"/regserver",False
Else
JoinParam exeParam,"/unregserver",False
End If
End If
ElseIf LCase(.GetExtensionName(Server))="dll" Then
Regsvr=.BuildPath(.GetSpecialFolder(1).Path,"regsvr32.exe")
If .FileExists(Regsvr)And .FileExists(Server)then
exeParam=""
JoinParam exeParam,Regsvr,False
JoinParam exeParam,"/s",False
If Not fRegister Then JoinParam exeParam,"/u",False
JoinParam exeParam,Server,False
End If
End If
End With
If exeParam<>"" Then
On Error Resume Next
With objWshShell.Exec(exeParam)
Do While .Status=0
WScript.Sleep 50
Loop
RegServer=(.ExitCode=0)
End With
On Error Goto 0
End If
End If
End Function
Dim objRegProvCache
Function RegProv
If Not IsObject(objRegProvCache)Then Set objRegProvCache=GetObject("winmgmts:\\.\root\default:StdRegProv")
Set RegProv=objRegProvCache
End Function
Function RegHKLM
RegHKLM=&H80000002
End Function
Function ExistsRegKey(ByVal RegHKEY,ByVal strKeyPath)
Dim arrSubKeys
ExistsRegKey=(RegProv.EnumKey(RegHKEY,strKeyPath,arrSubKeys)=0)
End Function
Function GetRegStringValue(ByVal RegHKEY,ByVal strKeyPath,ByVal strValueName)
Dim strValue
If RegProv.GetStringValue(RegHKEY,strKeyPath,strValueName,strValue)=0 Then
GetRegStringValue=strValue
Else
GetRegStringValue=Null
End If
End Function
Function StrToInt(ByVal str)
Dim val,tmp
val=0
On Error Resume Next
If Len(str)>2 Then If Left(str,2)="0x" Then str="&H"&Mid(str,3)
tmp=CCur(str)
If Err.Number=0 Then
If tmp>2147483647 Then
If tmp<=4294967295 Then val=CLng(tmp-4294967296)
Else
If tmp>=-2147483648 Then val=CLng(tmp)
End If
If Err.Number<>0 Then val=0
End If
On Error Goto 0
StrToInt=val
End Function
Function CompareVersion(ByVal verA,ByVal verB)
Dim arrA,arrB,num,flg,a,b
arrA=Split(verA,".")
arrB=Split(verB,".")
num=0
flg=0
Do
a=StrToInt(arrA(num))
b=StrToInt(arrB(num))
If a>b Then
flg=1
ElseIf a<b Then
flg=-1
End If
num=num+1
Loop While(flg=0)And(num<=UBound(arrA))And(num<=UBound(arrB))
CompareVersion=flg
End Function
Function CompareFirmwareVersion(ByVal verA,ByVal verB)
Dim arrA,arrB,num,flg,a,b
If LCase(Left(verA,1))="v" Then verA=Mid(verA,2)
If LCase(Left(verB,1))="v" Then verB=Mid(verB,2)
arrA=Split(verA,".")
arrB=Split(verB,".")
flg=0
a=arrA(0)
b=arrB(0)
If a>b Then
flg=1
ElseIf a<b Then
flg=-1
ElseIf(UBound(arrA)>0)And(UBound(arrB)>0)Then
verA=Left(arrA(1)&"   ",3)
verB=Left(arrB(1)&"   ",3)
a=Left(verA,1)
b=Left(verB,1)
If a>b Then
flg=1
ElseIf a<b Then
flg=-1
Else
a=Mid(verA,2,1)
b=Mid(verB,2,1)
If a>b Then
flg=1
ElseIf a<b Then
flg=-1
Else
a=Mid(verA,3,1)
b=Mid(verB,3,1)
If a>b Then
flg=1
ElseIf a<b Then
flg=-1
End If
End If
End If
End If
CompareFirmwareVersion=flg
End Function
