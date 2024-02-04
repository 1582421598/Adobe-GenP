#NoTrayIcon
#RequireAdmin
Global Const $STR_ENTIRESPLIT = 1
Global Const $STR_NOCOUNT = 2
Global Const $STR_REGEXPARRAYGLOBALFULLMATCH = 4
Func _StringRepeat($sString, $iRepeatCount)
$iRepeatCount = Int($iRepeatCount)
If $iRepeatCount = 0 Then Return ""
If StringLen($sString) < 1 Or $iRepeatCount < 0 Then Return SetError(1, 0, "")
Local $sResult = ""
While $iRepeatCount > 1
If BitAND($iRepeatCount, 1) Then $sResult &= $sString
$sString &= $sString
$iRepeatCount = BitShift($iRepeatCount, 1)
WEnd
Return $sString & $sResult
EndFunc
Global Const $PBS_SMOOTHREVERSE = 0x10
Global Const $WS_MAXIMIZEBOX = 0x00010000
Global Const $WS_MINIMIZEBOX = 0x00020000
Global Const $WS_SIZEBOX = 0x00040000
Global Const $WS_SYSMENU = 0x00080000
Global Const $WS_VSCROLL = 0x00200000
Global Const $WS_CAPTION = 0x00C00000
Global Const $WS_DISABLED = 0x08000000
Global Const $WS_POPUP = 0x80000000
Global Const $WM_NOTIFY = 0x004E
Global Const $WM_COMMAND = 0x0111
Global Const $NM_FIRST = 0
Global Const $NM_CLICK = $NM_FIRST - 2
Global Const $NM_RCLICK = $NM_FIRST - 5
Global Const $GUI_SS_DEFAULT_GUI = BitOR($WS_MINIMIZEBOX, $WS_CAPTION, $WS_POPUP, $WS_SYSMENU)
Global Const $GUI_EVENT_CLOSE = -3
Global Const $GUI_EVENT_RESTORE = -5
Global Const $GUI_EVENT_MAXIMIZE = -6
Global Const $GUI_EVENT_RESIZED = -12
Global Const $GUI_RUNDEFMSG = 'GUI_RUNDEFMSG'
Global Const $GUI_SHOW = 16
Global Const $GUI_HIDE = 32
Global Const $GUI_DISABLE = 128
Global Const $GUI_DOCKAUTO = 0x0001
Global Const $GUI_DOCKBOTTOM = 0x0040
Global Const $GUI_DOCKVCENTER = 0x0080
Global Const $ES_CENTER = 1
Global Const $ES_AUTOVSCROLL = 64
Global Const $ES_READONLY = 2048
Global Const $STDOUT_CHILD = 2
Global Const $STDERR_CHILD = 4
Global Const $UBOUND_DIMENSIONS = 0
Global Const $UBOUND_ROWS = 1
Global Const $UBOUND_COLUMNS = 2
Global Const $DIR_EXTENDED = 1
Global Const $MB_SYSTEMMODAL = 4096
Global Enum $ARRAYFILL_FORCE_DEFAULT, $ARRAYFILL_FORCE_SINGLEITEM, $ARRAYFILL_FORCE_INT, $ARRAYFILL_FORCE_NUMBER, $ARRAYFILL_FORCE_PTR, $ARRAYFILL_FORCE_HWND, $ARRAYFILL_FORCE_STRING, $ARRAYFILL_FORCE_BOOLEAN
Func _ArrayAdd(ByRef $aArray, $vValue, $iStart = 0, $sDelim_Item = "|", $sDelim_Row = @CRLF, $iForce = $ARRAYFILL_FORCE_DEFAULT)
If $iStart = Default Then $iStart = 0
If $sDelim_Item = Default Then $sDelim_Item = "|"
If $sDelim_Row = Default Then $sDelim_Row = @CRLF
If $iForce = Default Then $iForce = $ARRAYFILL_FORCE_DEFAULT
If Not IsArray($aArray) Then Return SetError(1, 0, -1)
Local $iDim_1 = UBound($aArray, $UBOUND_ROWS)
Local $hDataType = 0
Switch $iForce
Case $ARRAYFILL_FORCE_INT
$hDataType = Int
Case $ARRAYFILL_FORCE_NUMBER
$hDataType = Number
Case $ARRAYFILL_FORCE_PTR
$hDataType = Ptr
Case $ARRAYFILL_FORCE_HWND
$hDataType = Hwnd
Case $ARRAYFILL_FORCE_STRING
$hDataType = String
Case $ARRAYFILL_FORCE_BOOLEAN
$hDataType = "Boolean"
EndSwitch
Switch UBound($aArray, $UBOUND_DIMENSIONS)
Case 1
If $iForce = $ARRAYFILL_FORCE_SINGLEITEM Then
ReDim $aArray[$iDim_1 + 1]
$aArray[$iDim_1] = $vValue
Return $iDim_1
EndIf
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 1 Then Return SetError(5, 0, -1)
$hDataType = 0
Else
Local $aTmp = StringSplit($vValue, $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
If UBound($aTmp, $UBOUND_ROWS) = 1 Then
$aTmp[0] = $vValue
EndIf
$vValue = $aTmp
EndIf
Local $iAdd = UBound($vValue, $UBOUND_ROWS)
ReDim $aArray[$iDim_1 + $iAdd]
For $i = 0 To $iAdd - 1
If String($hDataType) = "Boolean" Then
Switch $vValue[$i]
Case "True", "1"
$aArray[$iDim_1 + $i] = True
Case "False", "0", ""
$aArray[$iDim_1 + $i] = False
EndSwitch
ElseIf IsFunc($hDataType) Then
$aArray[$iDim_1 + $i] = $hDataType($vValue[$i])
Else
$aArray[$iDim_1 + $i] = $vValue[$i]
EndIf
Next
Return $iDim_1 + $iAdd - 1
Case 2
Local $iDim_2 = UBound($aArray, $UBOUND_COLUMNS)
If $iStart < 0 Or $iStart > $iDim_2 - 1 Then Return SetError(4, 0, -1)
Local $iValDim_1, $iValDim_2 = 0, $iColCount
If IsArray($vValue) Then
If UBound($vValue, $UBOUND_DIMENSIONS) <> 2 Then Return SetError(5, 0, -1)
$iValDim_1 = UBound($vValue, $UBOUND_ROWS)
$iValDim_2 = UBound($vValue, $UBOUND_COLUMNS)
$hDataType = 0
Else
Local $aSplit_1 = StringSplit($vValue, $sDelim_Row, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iValDim_1 = UBound($aSplit_1, $UBOUND_ROWS)
Local $aTmp[$iValDim_1][0], $aSplit_2
For $i = 0 To $iValDim_1 - 1
$aSplit_2 = StringSplit($aSplit_1[$i], $sDelim_Item, $STR_NOCOUNT + $STR_ENTIRESPLIT)
$iColCount = UBound($aSplit_2)
If $iColCount > $iValDim_2 Then
$iValDim_2 = $iColCount
ReDim $aTmp[$iValDim_1][$iValDim_2]
EndIf
For $j = 0 To $iColCount - 1
$aTmp[$i][$j] = $aSplit_2[$j]
Next
Next
$vValue = $aTmp
EndIf
If UBound($vValue, $UBOUND_COLUMNS) + $iStart > UBound($aArray, $UBOUND_COLUMNS) Then Return SetError(3, 0, -1)
ReDim $aArray[$iDim_1 + $iValDim_1][$iDim_2]
For $iWriteTo_Index = 0 To $iValDim_1 - 1
For $j = 0 To $iDim_2 - 1
If $j < $iStart Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
ElseIf $j - $iStart > $iValDim_2 - 1 Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = ""
Else
If String($hDataType) = "Boolean" Then
Switch $vValue[$iWriteTo_Index][$j - $iStart]
Case "True", "1"
$aArray[$iWriteTo_Index + $iDim_1][$j] = True
Case "False", "0", ""
$aArray[$iWriteTo_Index + $iDim_1][$j] = False
EndSwitch
ElseIf IsFunc($hDataType) Then
$aArray[$iWriteTo_Index + $iDim_1][$j] = $hDataType($vValue[$iWriteTo_Index][$j - $iStart])
Else
$aArray[$iWriteTo_Index + $iDim_1][$j] = $vValue[$iWriteTo_Index][$j - $iStart]
EndIf
EndIf
Next
Next
Case Else
Return SetError(2, 0, -1)
EndSwitch
Return UBound($aArray, $UBOUND_ROWS) - 1
EndFunc
Global Const $MEM_COMMIT = 0x00001000
Global Const $MEM_RESERVE = 0x00002000
Global Const $PAGE_READWRITE = 0x00000004
Global Const $MEM_RELEASE = 0x00008000
Global Const $PROCESS_VM_OPERATION = 0x00000008
Global Const $PROCESS_VM_READ = 0x00000010
Global Const $PROCESS_VM_WRITE = 0x00000020
Global Const $SE_DEBUG_NAME = "SeDebugPrivilege"
Global Const $SE_PRIVILEGE_ENABLED = 0x00000002
Global Enum $SECURITYANONYMOUS = 0, $SECURITYIDENTIFICATION, $SECURITYIMPERSONATION, $SECURITYDELEGATION
Global Const $TOKEN_QUERY = 0x00000008
Global Const $TOKEN_ADJUST_PRIVILEGES = 0x00000020
Func _WinAPI_GetLastError(Const $_iCallerError = @error, Const $_iCallerExtended = @extended)
Local $aCall = DllCall("kernel32.dll", "dword", "GetLastError")
Return SetError($_iCallerError, $_iCallerExtended, $aCall[0])
EndFunc
Func _Security__AdjustTokenPrivileges($hToken, $bDisableAll, $tNewState, $iBufferLen, $tPrevState = 0, $pRequired = 0)
Local $aCall = DllCall("advapi32.dll", "bool", "AdjustTokenPrivileges", "handle", $hToken, "bool", $bDisableAll, "struct*", $tNewState, "dword", $iBufferLen, "struct*", $tPrevState, "struct*", $pRequired)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__ImpersonateSelf($iLevel = $SECURITYIMPERSONATION)
Local $aCall = DllCall("advapi32.dll", "bool", "ImpersonateSelf", "int", $iLevel)
If @error Then Return SetError(@error, @extended, False)
Return Not($aCall[0] = 0)
EndFunc
Func _Security__LookupPrivilegeValue($sSystem, $sName)
Local $aCall = DllCall("advapi32.dll", "bool", "LookupPrivilegeValueW", "wstr", $sSystem, "wstr", $sName, "int64*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)
Return $aCall[3]
EndFunc
Func _Security__OpenThreadToken($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $aCall
If $hThread = 0 Then
$aCall = DllCall("kernel32.dll", "handle", "GetCurrentThread")
If @error Then Return SetError(@error + 20, @extended, 0)
$hThread = $aCall[0]
EndIf
$aCall = DllCall("advapi32.dll", "bool", "OpenThreadToken", "handle", $hThread, "dword", $iAccess, "bool", $bOpenAsSelf, "handle*", 0)
If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)
Return $aCall[4]
EndFunc
Func _Security__OpenThreadTokenEx($iAccess, $hThread = 0, $bOpenAsSelf = False)
Local $hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then
Local Const $ERROR_NO_TOKEN = 1008
If _WinAPI_GetLastError() <> $ERROR_NO_TOKEN Then Return SetError(20, _WinAPI_GetLastError(), 0)
If Not _Security__ImpersonateSelf() Then Return SetError(@error + 10, _WinAPI_GetLastError(), 0)
$hToken = _Security__OpenThreadToken($iAccess, $hThread, $bOpenAsSelf)
If $hToken = 0 Then Return SetError(@error, _WinAPI_GetLastError(), 0)
EndIf
Return $hToken
EndFunc
Func _Security__SetPrivilege($hToken, $sPrivilege, $bEnable)
Local $iLUID = _Security__LookupPrivilegeValue("", $sPrivilege)
If $iLUID = 0 Then Return SetError(@error + 10, @extended, False)
Local Const $tagTOKEN_PRIVILEGES = "dword Count;align 4;int64 LUID;dword Attributes"
Local $tCurrState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iCurrState = DllStructGetSize($tCurrState)
Local $tPrevState = DllStructCreate($tagTOKEN_PRIVILEGES)
Local $iPrevState = DllStructGetSize($tPrevState)
Local $tRequired = DllStructCreate("int Data")
DllStructSetData($tCurrState, "Count", 1)
DllStructSetData($tCurrState, "LUID", $iLUID)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tCurrState, $iCurrState, $tPrevState, $tRequired) Then Return SetError(2, @error, False)
DllStructSetData($tPrevState, "Count", 1)
DllStructSetData($tPrevState, "LUID", $iLUID)
Local $iAttributes = DllStructGetData($tPrevState, "Attributes")
If $bEnable Then
$iAttributes = BitOR($iAttributes, $SE_PRIVILEGE_ENABLED)
Else
$iAttributes = BitAND($iAttributes, BitNOT($SE_PRIVILEGE_ENABLED))
EndIf
DllStructSetData($tPrevState, "Attributes", $iAttributes)
If Not _Security__AdjustTokenPrivileges($hToken, False, $tPrevState, $iPrevState, $tCurrState, $tRequired) Then Return SetError(3, @error, False)
Return True
EndFunc
Global Const $tagPOINT = "struct;long X;long Y;endstruct"
Global Const $tagRECT = "struct;long Left;long Top;long Right;long Bottom;endstruct"
Global Const $tagNMHDR = "struct;hwnd hWndFrom;uint_ptr IDFrom;INT Code;endstruct"
Global Const $tagLVHITTESTINFO = $tagPOINT & ";uint Flags;int Item;int SubItem;int iGroup"
Global Const $tagLVITEM = "struct;uint Mask;int Item;int SubItem;uint State;uint StateMask;ptr Text;int TextMax;int Image;lparam Param;" & "int Indent;int GroupID;uint Columns;ptr pColumns;ptr piColFmt;int iGroup;endstruct"
Global Const $tagNMITEMACTIVATE = $tagNMHDR & ";int Index;int SubItem;uint NewState;uint OldState;uint Changed;" & $tagPOINT & ";lparam lParam;uint KeyFlags"
Global Const $tagSECURITY_ATTRIBUTES = "dword Length;ptr Descriptor;bool InheritHandle"
Global Const $tagMEMMAP = "handle hProc;ulong_ptr Size;ptr Mem"
Func _MemFree(ByRef $tMemMap)
Local $pMemory = DllStructGetData($tMemMap, "Mem")
Local $hProcess = DllStructGetData($tMemMap, "hProc")
Local $bResult = _MemVirtualFreeEx($hProcess, $pMemory, 0, $MEM_RELEASE)
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hProcess)
If @error Then Return SetError(@error, @extended, False)
Return $bResult
EndFunc
Func _MemInit($hWnd, $iSize, ByRef $tMemMap)
Local $aCall = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error + 10, @extended, 0)
Local $iProcessID = $aCall[2]
If $iProcessID = 0 Then Return SetError(1, 0, 0)
Local $iAccess = BitOR($PROCESS_VM_OPERATION, $PROCESS_VM_READ, $PROCESS_VM_WRITE)
Local $hProcess = __Mem_OpenProcess($iAccess, False, $iProcessID, True)
Local $iAlloc = BitOR($MEM_RESERVE, $MEM_COMMIT)
Local $pMemory = _MemVirtualAllocEx($hProcess, 0, $iSize, $iAlloc, $PAGE_READWRITE)
If $pMemory = 0 Then Return SetError(2, 0, 0)
$tMemMap = DllStructCreate($tagMEMMAP)
DllStructSetData($tMemMap, "hProc", $hProcess)
DllStructSetData($tMemMap, "Size", $iSize)
DllStructSetData($tMemMap, "Mem", $pMemory)
Return $pMemory
EndFunc
Func _MemRead(ByRef $tMemMap, $pSrce, $pDest, $iSize)
Local $aCall = DllCall("kernel32.dll", "bool", "ReadProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pSrce, "struct*", $pDest, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Func _MemWrite(ByRef $tMemMap, $pSrce, $pDest = 0, $iSize = 0, $sSrce = "struct*")
If $pDest = 0 Then $pDest = DllStructGetData($tMemMap, "Mem")
If $iSize = 0 Then $iSize = DllStructGetData($tMemMap, "Size")
Local $aCall = DllCall("kernel32.dll", "bool", "WriteProcessMemory", "handle", DllStructGetData($tMemMap, "hProc"), "ptr", $pDest, $sSrce, $pSrce, "ulong_ptr", $iSize, "ulong_ptr*", 0)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Func _MemVirtualAllocEx($hProcess, $pAddress, $iSize, $iAllocation, $iProtect)
Local $aCall = DllCall("kernel32.dll", "ptr", "VirtualAllocEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iAllocation, "dword", $iProtect)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Func _MemVirtualFreeEx($hProcess, $pAddress, $iSize, $iFreeType)
Local $aCall = DllCall("kernel32.dll", "bool", "VirtualFreeEx", "handle", $hProcess, "ptr", $pAddress, "ulong_ptr", $iSize, "dword", $iFreeType)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Func __Mem_OpenProcess($iAccess, $bInherit, $iPID, $bDebugPriv = False)
Local $aCall = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
If @error Then Return SetError(@error, @extended, 0)
If $aCall[0] Then Return $aCall[0]
If Not $bDebugPriv Then Return SetError(100, 0, 0)
Local $hToken = _Security__OpenThreadTokenEx(BitOR($TOKEN_ADJUST_PRIVILEGES, $TOKEN_QUERY))
If @error Then Return SetError(@error + 10, @extended, 0)
_Security__SetPrivilege($hToken, $SE_DEBUG_NAME, True)
Local $iError = @error
Local $iExtended = @extended
Local $iRet = 0
If Not @error Then
$aCall = DllCall("kernel32.dll", "handle", "OpenProcess", "dword", $iAccess, "bool", $bInherit, "dword", $iPID)
$iError = @error
$iExtended = @extended
If $aCall[0] Then $iRet = $aCall[0]
_Security__SetPrivilege($hToken, $SE_DEBUG_NAME, False)
If @error Then
$iError = @error + 20
$iExtended = @extended
EndIf
Else
$iError = @error + 30
EndIf
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $hToken)
Return SetError($iError, $iExtended, $iRet)
EndFunc
Func _SendMessage($hWnd, $iMsg, $wParam = 0, $lParam = 0, $iReturn = 0, $wParamType = "wparam", $lParamType = "lparam", $sReturnType = "lresult")
Local $aCall = DllCall("user32.dll", $sReturnType, "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, $wParamType, $wParam, $lParamType, $lParam)
If @error Then Return SetError(@error, @extended, "")
If $iReturn >= 0 And $iReturn <= 4 Then Return $aCall[$iReturn]
Return $aCall
EndFunc
Global Const $FC_OVERWRITE = 1
Global Const $FO_READ = 0
Global Const $FO_OVERWRITE = 2
Global Const $FO_BINARY = 16
Func _WinAPI_GetDlgCtrlID($hWnd)
Local $aCall = DllCall("user32.dll", "int", "GetDlgCtrlID", "hwnd", $hWnd)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Func _WinAPI_GetString($pString, $bUnicode = True)
Local $iLength = _WinAPI_StrLen($pString, $bUnicode)
If @error Or Not $iLength Then Return SetError(@error + 10, @extended, '')
Local $tString = DllStructCreate(($bUnicode ? 'wchar' : 'char') & '[' &($iLength + 1) & ']', $pString)
If @error Then Return SetError(@error, @extended, '')
Return SetExtended($iLength, DllStructGetData($tString, 1))
EndFunc
Func _WinAPI_StrLen($pString, $bUnicode = True)
Local $W = ''
If $bUnicode Then $W = 'W'
Local $aCall = DllCall('kernel32.dll', 'int', 'lstrlen' & $W, 'struct*', $pString)
If @error Then Return SetError(@error, @extended, 0)
Return $aCall[0]
EndFunc
Global $__g_aInProcess_WinAPI[64][2] = [[0, 0]]
Func _WinAPI_GetWindowThreadProcessId($hWnd, ByRef $iPID)
Local $aCall = DllCall("user32.dll", "dword", "GetWindowThreadProcessId", "hwnd", $hWnd, "dword*", 0)
If @error Then Return SetError(@error, @extended, 0)
$iPID = $aCall[2]
Return $aCall[0]
EndFunc
Func _WinAPI_InProcess($hWnd, ByRef $hLastWnd)
If $hWnd = $hLastWnd Then Return True
For $iI = $__g_aInProcess_WinAPI[0][0] To 1 Step -1
If $hWnd = $__g_aInProcess_WinAPI[$iI][0] Then
If $__g_aInProcess_WinAPI[$iI][1] Then
$hLastWnd = $hWnd
Return True
Else
Return False
EndIf
EndIf
Next
Local $iPID
_WinAPI_GetWindowThreadProcessId($hWnd, $iPID)
Local $iCount = $__g_aInProcess_WinAPI[0][0] + 1
If $iCount >= 64 Then $iCount = 1
$__g_aInProcess_WinAPI[0][0] = $iCount
$__g_aInProcess_WinAPI[$iCount][0] = $hWnd
$__g_aInProcess_WinAPI[$iCount][1] =($iPID = @AutoItPID)
Return $__g_aInProcess_WinAPI[$iCount][1]
EndFunc
Func _WinAPI_InvalidateRect($hWnd, $tRECT = 0, $bErase = True)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $aCall = DllCall("user32.dll", "bool", "InvalidateRect", "hwnd", $hWnd, "struct*", $tRECT, "bool", $bErase)
If @error Then Return SetError(@error, @extended, False)
Return $aCall[0]
EndFunc
Global $__g_hGUICtrl_LastWnd
Func __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
If $iElement > 0 Then
DllStructSetData($tItem, $iElement, DllStructGetPtr($tBuffer))
If $iElement = $iElementMax Then DllStructSetData($tItem, $iElement + 1, DllStructGetSize($tBuffer))
EndIf
Local $iRet
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hGUICtrl_LastWnd) Then
$iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "struct*", $tItem)[0]
Else
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap, $pText
Local $iBuffer = 0
If($iElement > 0) Or($iElementMax = 0) Then $iBuffer = DllStructGetSize($tBuffer)
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
If $iBuffer Then
$pText = $pMemory + $iItem
If $iElementMax Then
DllStructSetData($tItem, $iElement, $pText)
Else
$iIndex = $pText
EndIf
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
EndIf
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
$iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "ptr", $pMemory)[0]
If $iBuffer And $bRetBuffer Then _MemRead($tMemMap, $pText, $tBuffer, $iBuffer)
If $bRetItem Then _MemRead($tMemMap, $pMemory, $tItem, $iItem)
_MemFree($tMemMap)
EndIf
Else
$iRet = GUICtrlSendMsg($hWnd, $iMsg, $iIndex, DllStructGetPtr($tItem))
EndIf
Return $iRet
EndFunc
Func __GUICtrl_SendMsg_Init($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
#forceref $iMsg, $iIndex, $bRetItem, $bRetBuffer
DllStructSetData($tItem, $iElement, DllStructGetPtr($tBuffer))
If $iElement = $iElementMax Then DllStructSetData($tItem, $iElement + 1, DllStructGetSize($tBuffer))
Local $pFunc
If IsHWnd($hWnd) Then
If _WinAPI_InProcess($hWnd, $__g_hGUICtrl_LastWnd) Then
$pFunc = __GUICtrl_SendMsg_InProcess
SetExtended(1)
Else
$pFunc = __GUICtrl_SendMsg_OutProcess
SetExtended(2)
EndIf
Else
$pFunc = __GUICtrl_SendMsg_Internal
SetExtended(3)
EndIf
Return $pFunc
EndFunc
Func __GUICtrl_SendMsg_InProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
#forceref $tBuffer, $bRetItem, $bRetBuffer, $iElementMax
Return DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "struct*", $tItem)[0]
EndFunc
Func __GUICtrl_SendMsg_OutProcess($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
Local $iItem = DllStructGetSize($tItem)
Local $tMemMap, $pText
Local $iBuffer = 0
If($iElement > 0) Or($iElementMax = 0) Then $iBuffer = DllStructGetSize($tBuffer)
Local $pMemory = _MemInit($hWnd, $iItem + $iBuffer, $tMemMap)
If $iBuffer Then
$pText = $pMemory + $iItem
If $iElementMax Then
DllStructSetData($tItem, $iElement, $pText)
Else
$iIndex = $pText
EndIf
_MemWrite($tMemMap, $tBuffer, $pText, $iBuffer)
EndIf
_MemWrite($tMemMap, $tItem, $pMemory, $iItem)
Local $iRet = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", $iMsg, "wparam", $iIndex, "ptr", $pMemory)[0]
If $iBuffer And $bRetBuffer Then _MemRead($tMemMap, $pText, $tBuffer, $iBuffer)
If $bRetItem Then _MemRead($tMemMap, $pMemory, $tItem, $iItem)
_MemFree($tMemMap)
Return $iRet
EndFunc
Func __GUICtrl_SendMsg_Internal($hWnd, $iMsg, $iIndex, ByRef $tItem, $tBuffer = 0, $bRetItem = False, $iElement = -1, $bRetBuffer = False, $iElementMax = $iElement)
#forceref $tBuffer, $bRetItem, $bRetBuffer, $iElementMax
Return GUICtrlSendMsg($hWnd, $iMsg, $iIndex, DllStructGetPtr($tItem))
EndFunc
Global Const $_UDF_STARTID = 10000
Func _WinAPI_MultiByteToWideChar($vText, $iCodePage = 0, $iFlags = 0, $bRetString = False)
Local $sTextType = ""
If IsString($vText) Then $sTextType = "str"
If(IsDllStruct($vText) Or IsPtr($vText)) Then $sTextType = "struct*"
If $sTextType = "" Then Return SetError(1, 0, 0)
Local $aCall = DllCall("kernel32.dll", "int", "MultiByteToWideChar", "uint", $iCodePage, "dword", $iFlags, $sTextType, $vText, "int", -1, "ptr", 0, "int", 0)
If @error Or Not $aCall[0] Then Return SetError(@error + 10, @extended, 0)
Local $iOut = $aCall[0]
Local $tOut = DllStructCreate("wchar[" & $iOut & "]")
$aCall = DllCall("kernel32.dll", "int", "MultiByteToWideChar", "uint", $iCodePage, "dword", $iFlags, $sTextType, $vText, "int", -1, "struct*", $tOut, "int", $iOut)
If @error Or Not $aCall[0] Then Return SetError(@error + 20, @extended, 0)
If $bRetString Then Return DllStructGetData($tOut, 1)
Return $tOut
EndFunc
Global Const $LVGS_NORMAL = 0x00000000
Global Const $LVGS_COLLAPSED = 0x00000001
Global Const $LVGS_COLLAPSIBLE = 0x00000008
Global Const $LVGS_SELECTED = 0x00000020
Global Const $LV_ERR = -1
Global Const $LV_VIEW_DETAILS = 0x0001
Global Const $LV_VIEW_ICON = 0x0000
Global Const $LV_VIEW_LIST = 0x0003
Global Const $LV_VIEW_SMALLICON = 0x0002
Global Const $LV_VIEW_TILE = 0x0004
Global Const $LVCF_FMT = 0x0001
Global Const $LVCF_IMAGE = 0x0010
Global Const $LVCF_TEXT = 0x0004
Global Const $LVCF_WIDTH = 0x0002
Global Const $LVCFMT_BITMAP_ON_RIGHT = 0x1000
Global Const $LVCFMT_CENTER = 0x0002
Global Const $LVCFMT_COL_HAS_IMAGES = 0x8000
Global Const $LVCFMT_IMAGE = 0x0800
Global Const $LVCFMT_LEFT = 0x0000
Global Const $LVCFMT_RIGHT = 0x0001
Global Const $LVGA_HEADER_LEFT = 0x00000001
Global Const $LVGA_HEADER_CENTER = 0x00000002
Global Const $LVGA_HEADER_RIGHT = 0x00000004
Global Const $LVGF_ALIGN = 0x00000008
Global Const $LVGF_GROUPID = 0x00000010
Global Const $LVGF_HEADER = 0x00000001
Global Const $LVGF_ITEMS = 0x00004000
Global Const $LVGF_STATE = 0x00000004
Global Const $LVHT_ABOVE = 0x00000008
Global Const $LVHT_BELOW = 0x00000010
Global Const $LVHT_NOWHERE = 0x00000001
Global Const $LVHT_ONITEMICON = 0x00000002
Global Const $LVHT_ONITEMLABEL = 0x00000004
Global Const $LVHT_ONITEMSTATEICON = 0x00000008
Global Const $LVHT_TOLEFT = 0x00000040
Global Const $LVHT_TORIGHT = 0x00000020
Global Const $LVHT_ONITEM = BitOR($LVHT_ONITEMICON, $LVHT_ONITEMLABEL, $LVHT_ONITEMSTATEICON)
Global Const $LVIF_GROUPID = 0x00000100
Global Const $LVIF_IMAGE = 0x00000002
Global Const $LVIF_PARAM = 0x00000004
Global Const $LVIF_STATE = 0x00000008
Global Const $LVIF_TEXT = 0x00000001
Global Const $LVIS_FOCUSED = 0x0001
Global Const $LVIS_SELECTED = 0x0002
Global Const $LVS_EX_CHECKBOXES = 0x00000004
Global Const $LVS_EX_DOUBLEBUFFER = 0x00010000
Global Const $LVS_EX_FULLROWSELECT = 0x00000020
Global Const $LVS_EX_GRIDLINES = 0x00000001
Global Const $LVS_EX_SUBITEMIMAGES = 0x00000002
Global Const $LVM_FIRST = 0x1000
Global Const $LVM_DELETEALLITEMS =($LVM_FIRST + 9)
Global Const $LVM_ENABLEGROUPVIEW =($LVM_FIRST + 157)
Global Const $LVM_ENSUREVISIBLE =($LVM_FIRST + 19)
Global Const $LVM_GETGROUPCOUNT =($LVM_FIRST + 152)
Global Const $LVM_GETGROUPINFO =($LVM_FIRST + 149)
Global Const $LVM_GETGROUPINFOBYINDEX =($LVM_FIRST + 153)
Global Const $LVM_GETHEADER =($LVM_FIRST + 31)
Global Const $LVM_GETITEMA =($LVM_FIRST + 5)
Global Const $LVM_GETITEMW =($LVM_FIRST + 75)
Global Const $LVM_GETITEMCOUNT =($LVM_FIRST + 4)
Global Const $LVM_GETITEMRECT =($LVM_FIRST + 14)
Global Const $LVM_GETITEMTEXTA =($LVM_FIRST + 45)
Global Const $LVM_GETITEMTEXTW =($LVM_FIRST + 115)
Global Const $LVM_GETSELECTEDCOUNT =($LVM_FIRST + 50)
Global Const $LVM_GETUNICODEFORMAT = 0x2000 + 6
Global Const $LVM_GETVIEW =($LVM_FIRST + 143)
Global Const $LVM_GETVIEWRECT =($LVM_FIRST + 34)
Global Const $LVM_HITTEST =($LVM_FIRST + 18)
Global Const $LVM_INSERTCOLUMNA =($LVM_FIRST + 27)
Global Const $LVM_INSERTCOLUMNW =($LVM_FIRST + 97)
Global Const $LVM_INSERTGROUP =($LVM_FIRST + 145)
Global Const $LVM_INSERTITEMA =($LVM_FIRST + 7)
Global Const $LVM_INSERTITEMW =($LVM_FIRST + 77)
Global Const $LVM_REMOVEGROUP =($LVM_FIRST + 150)
Global Const $LVM_SCROLL =($LVM_FIRST + 20)
Global Const $LVM_SETCOLUMNA =($LVM_FIRST + 26)
Global Const $LVM_SETCOLUMNW =($LVM_FIRST + 96)
Global Const $LVM_SETCOLUMNWIDTH =($LVM_FIRST + 30)
Global Const $LVM_SETEXTENDEDLISTVIEWSTYLE =($LVM_FIRST + 54)
Global Const $LVM_SETGROUPINFO =($LVM_FIRST + 147)
Global Const $LVM_SETITEMA =($LVM_FIRST + 6)
Global Const $LVM_SETITEMW =($LVM_FIRST + 76)
Global Const $LVM_SETITEMCOUNT =($LVM_FIRST + 47)
Global Const $LVM_SETITEMSTATE =($LVM_FIRST + 43)
Global Const $LVN_FIRST = -100
Global Const $LVN_COLUMNCLICK =($LVN_FIRST - 8)
Global Const $LVSICF_NOINVALIDATEALL = 0x00000001
Global Const $LVSICF_NOSCROLL = 0x00000002
Global Const $__LISTVIEWCONSTANT_SORTINFOSIZE = 11
Global $__g_aListViewSortInfo[1][$__LISTVIEWCONSTANT_SORTINFOSIZE]
Global $__g_tListViewBuffer, $__g_tListViewBufferANSI
Global $__g_tListViewItem = DllStructCreate($tagLVITEM)
Global Const $__LISTVIEWCONSTANT_WM_SETREDRAW = 0x000B
Global Const $tagLVCOLUMN = "uint Mask;int Fmt;int CX;ptr Text;int TextMax;int SubItem;int Image;int Order;int cxMin;int cxDefault;int cxIdeal"
Global Const $tagLVGROUP = "uint Size;uint Mask;ptr Header;int HeaderMax;ptr Footer;int FooterMax;int GroupID;uint StateMask;uint State;uint Align;" & "ptr  pszSubtitle;uint cchSubtitle;ptr pszTask;uint cchTask;ptr pszDescriptionTop;uint cchDescriptionTop;ptr pszDescriptionBottom;" & "uint cchDescriptionBottom;int iTitleImage;int iExtendedImage;int iFirstItem;uint cItems;ptr pszSubsetTitle;uint cchSubsetTitle"
Func _GUICtrlListView_AddArray($hWnd, ByRef $aItems)
Local $tBuffer, $iMsg, $iMsgSet
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_INSERTITEMW
$iMsgSet = $LVM_SETITEMW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_INSERTITEMA
$iMsgSet = $LVM_SETITEMA
EndIf
Local $tItem = $__g_tListViewItem
DllStructSetData($tItem, "Mask", $LVIF_TEXT)
Local $iLastItem = _GUICtrlListView_GetItemCount($hWnd)
_GUICtrlListView_BeginUpdate($hWnd)
Local $pSendMsg = __GUICtrl_SendMsg_Init($hWnd, $iMsg, 0, $tItem, $tBuffer, False, 6)
For $iI = 0 To UBound($aItems) - 1
DllStructSetData($tItem, "Item", $iI + $iLastItem)
DllStructSetData($tItem, "SubItem", 0)
DllStructSetData($tBuffer, 1, $aItems[$iI][0])
$pSendMsg($hWnd, $iMsg, 0, $tItem, $tBuffer, False, 6)
For $iJ = 1 To UBound($aItems, $UBOUND_COLUMNS) - 1
DllStructSetData($tItem, "SubItem", $iJ)
DllStructSetData($tBuffer, 1, $aItems[$iI][$iJ])
$pSendMsg($hWnd, $iMsgSet, 0, $tItem, $tBuffer, False, 6)
Next
Next
_GUICtrlListView_EndUpdate($hWnd)
EndFunc
Func _GUICtrlListView_AddColumn($hWnd, $sText, $iWidth = 50, $iAlign = -1, $iImage = -1, $bOnRight = False)
Return _GUICtrlListView_InsertColumn($hWnd, _GUICtrlListView_GetColumnCount($hWnd), $sText, $iWidth, $iAlign, $iImage, $bOnRight)
EndFunc
Func _GUICtrlListView_AddItem($hWnd, $sText, $iImage = -1, $iParam = 0)
Return _GUICtrlListView_InsertItem($hWnd, $sText, -1, $iImage, $iParam)
EndFunc
Func _GUICtrlListView_AddSubItem($hWnd, $iIndex, $sText, $iSubItem, $iImage = -1)
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_SETITEMW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_SETITEMA
EndIf
Local $tItem = $__g_tListViewItem
Local $iMask = $LVIF_TEXT
If $iImage <> -1 Then $iMask = BitOR($iMask, $LVIF_IMAGE)
DllStructSetData($tBuffer, 1, $sText)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
DllStructSetData($tItem, "Image", $iImage)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, $tBuffer, False, 6, False, -1)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_BeginUpdate($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, False) = 0
Else
Return GUICtrlSendMsg($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, False, 0) = 0
EndIf
EndFunc
Func _GUICtrlListView_DeleteAllItems($hWnd)
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
Local $vCID = 0
If IsHWnd($hWnd) Then
$vCID = _WinAPI_GetDlgCtrlID($hWnd)
Else
$vCID = $hWnd
$hWnd = GUICtrlGetHandle($hWnd)
EndIf
If $vCID < $_UDF_STARTID Then
Local $iParam = 0
For $iIndex = _GUICtrlListView_GetItemCount($hWnd) - 1 To 0 Step -1
$iParam = _GUICtrlListView_GetItemParam($hWnd, $iIndex)
If GUICtrlGetState($iParam) > 0 And GUICtrlGetHandle($iParam) = 0 Then
GUICtrlDelete($iParam)
EndIf
Next
If _GUICtrlListView_GetItemCount($hWnd) = 0 Then Return True
EndIf
Return _SendMessage($hWnd, $LVM_DELETEALLITEMS) <> 0
EndFunc
Func _GUICtrlListView_EnableGroupView($hWnd, $bEnable = True)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_ENABLEGROUPVIEW, $bEnable)
Else
Return GUICtrlSendMsg($hWnd, $LVM_ENABLEGROUPVIEW, $bEnable, 0)
EndIf
EndFunc
Func _GUICtrlListView_EndUpdate($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, True) = 0
Else
Return GUICtrlSendMsg($hWnd, $__LISTVIEWCONSTANT_WM_SETREDRAW, True, 0) = 0
EndIf
EndFunc
Func _GUICtrlListView_EnsureVisible($hWnd, $iIndex, $bPartialOK = False)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_ENSUREVISIBLE, $iIndex, $bPartialOK)
Else
Return GUICtrlSendMsg($hWnd, $LVM_ENSUREVISIBLE, $iIndex, $bPartialOK)
EndIf
EndFunc
Func _GUICtrlListView_GetColumnCount($hWnd)
Return _SendMessage(_GUICtrlListView_GetHeader($hWnd), 0x1200)
EndFunc
Func _GUICtrlListView_GetGroupCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETGROUPCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETGROUPCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetGroupInfo($hWnd, $iGroupID)
Local $tGroup = __GUICtrlListView_GetGroupInfoEx($hWnd, $iGroupID, BitOR($LVGF_HEADER, $LVGF_ALIGN))
Local $iErr = @error
Local $aGroup[2]
$aGroup[0] = _WinAPI_GetString(DllStructGetData($tGroup, "Header"))
Select
Case BitAND(DllStructGetData($tGroup, "Align"), $LVGA_HEADER_CENTER) <> 0
$aGroup[1] = 1
Case BitAND(DllStructGetData($tGroup, "Align"), $LVGA_HEADER_RIGHT) <> 0
$aGroup[1] = 2
Case Else
$aGroup[1] = 0
EndSelect
Return SetError($iErr, 0, $aGroup)
EndFunc
Func __GUICtrlListView_GetGroupInfoEx($hWnd, $iGroupID, $iMask)
Local $tGroup = DllStructCreate($tagLVGROUP)
Local $iGroup = DllStructGetSize($tGroup)
DllStructSetData($tGroup, "Size", $iGroup)
DllStructSetData($tGroup, "Mask", $iMask)
Local $iRet = __GUICtrl_SendMsg($hWnd, $LVM_GETGROUPINFO, $iGroupID, $tGroup, 0, True, -1)
Return SetError($iRet <> $iGroupID, 0, $tGroup)
EndFunc
Func _GUICtrlListView_GetGroupInfoByIndex($hWnd, $iIndex)
Local $tGroup = DllStructCreate($tagLVGROUP)
Local $iGroup = DllStructGetSize($tGroup)
DllStructSetData($tGroup, "Size", $iGroup)
DllStructSetData($tGroup, "Mask", BitOR($LVGF_HEADER, $LVGF_ALIGN, $LVGF_GROUPID))
Local $iRet = __GUICtrl_SendMsg($hWnd, $LVM_GETGROUPINFOBYINDEX, $iIndex, $tGroup, 0, True, -1)
Local $aGroup[3]
$aGroup[0] = _WinAPI_GetString(DllStructGetData($tGroup, "Header"))
Select
Case BitAND(DllStructGetData($tGroup, "Align"), $LVGA_HEADER_CENTER) <> 0
$aGroup[1] = 1
Case BitAND(DllStructGetData($tGroup, "Align"), $LVGA_HEADER_RIGHT) <> 0
$aGroup[1] = 2
Case Else
$aGroup[1] = 0
EndSelect
$aGroup[2] = DllStructGetData($tGroup, "GroupID")
Return SetError($iRet = 0, 0, $aGroup)
EndFunc
Func _GUICtrlListView_GetHeader($hWnd)
If IsHWnd($hWnd) Then
Return HWnd(_SendMessage($hWnd, $LVM_GETHEADER))
Else
Return HWnd(GUICtrlSendMsg($hWnd, $LVM_GETHEADER, 0, 0))
EndIf
EndFunc
Func _GUICtrlListView_GetItemChecked($hWnd, $iIndex)
Local $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$iMsg = $LVM_GETITEMW
Else
$iMsg = $LVM_GETITEMA
EndIf
Local $tItem = $__g_tListViewItem
DllStructSetData($tItem, "Mask", $LVIF_STATE)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
DllStructSetData($tItem, "StateMask", 0xffff)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, 0, True, -1)
If Not $iRet Then Return SetError($LV_ERR, $LV_ERR, False)
Return BitAND(DllStructGetData($tItem, "State"), 0x2000) <> 0
EndFunc
Func _GUICtrlListView_GetItemCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETITEMCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETITEMCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetItemEx($hWnd, ByRef $tItem)
Local $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$iMsg = $LVM_GETITEMW
Else
$iMsg = $LVM_GETITEMA
EndIf
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, 0, True, -1)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_GetItemGroupID($hWnd, $iIndex)
Local $tItem = $__g_tListViewItem
DllStructSetData($tItem, "Mask", $LVIF_GROUPID)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Return DllStructGetData($tItem, "GroupID")
EndFunc
Func _GUICtrlListView_GetItemParam($hWnd, $iIndex)
Local $tItem = $__g_tListViewItem
DllStructSetData($tItem, "Mask", $LVIF_PARAM)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
_GUICtrlListView_GetItemEx($hWnd, $tItem)
Return DllStructGetData($tItem, "Param")
EndFunc
Func _GUICtrlListView_GetItemRect($hWnd, $iIndex, $iPart = 3)
Local $tRECT = _GUICtrlListView_GetItemRectEx($hWnd, $iIndex, $iPart)
Local $aRect[4]
$aRect[0] = DllStructGetData($tRECT, "Left")
$aRect[1] = DllStructGetData($tRECT, "Top")
$aRect[2] = DllStructGetData($tRECT, "Right")
$aRect[3] = DllStructGetData($tRECT, "Bottom")
Return $aRect
EndFunc
Func _GUICtrlListView_GetItemRectEx($hWnd, $iIndex, $iPart = 3)
Local $tRECT = DllStructCreate($tagRECT)
DllStructSetData($tRECT, "Left", $iPart)
__GUICtrl_SendMsg($hWnd, $LVM_GETITEMRECT, $iIndex, $tRECT, 0, True, -1)
Return $tRECT
EndFunc
Func _GUICtrlListView_GetItemText($hWnd, $iIndex, $iSubItem = 0)
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_GETITEMTEXTW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_GETITEMTEXTA
EndIf
Local $tItem = $__g_tListViewItem
DllStructSetData($tBuffer, 1, "")
DllStructSetData($tItem, "Mask", $LVIF_TEXT)
DllStructSetData($tItem, "SubItem", $iSubItem)
__GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, $tItem, $tBuffer, False, 6, True)
Return DllStructGetData($tBuffer, 1)
EndFunc
Func _GUICtrlListView_GetSelectedCount($hWnd)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETSELECTEDCOUNT)
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETSELECTEDCOUNT, 0, 0)
EndIf
EndFunc
Func _GUICtrlListView_GetUnicodeFormat($hWnd)
If Not IsDllStruct($__g_tListViewBuffer) Then
$__g_tListViewBuffer = DllStructCreate("wchar Text[4096]")
$__g_tListViewBufferANSI = DllStructCreate("char Text[4096]", DllStructGetPtr($__g_tListViewBuffer))
EndIf
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_GETUNICODEFORMAT) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_GETUNICODEFORMAT, 0, 0) <> 0
EndIf
EndFunc
Func _GUICtrlListView_GetView($hWnd)
Local $iView
If IsHWnd($hWnd) Then
$iView = _SendMessage($hWnd, $LVM_GETVIEW)
Else
$iView = GUICtrlSendMsg($hWnd, $LVM_GETVIEW, 0, 0)
EndIf
Switch $iView
Case $LV_VIEW_ICON
Return Int($LV_VIEW_ICON)
Case $LV_VIEW_DETAILS
Return Int($LV_VIEW_DETAILS)
Case $LV_VIEW_LIST
Return Int($LV_VIEW_LIST)
Case $LV_VIEW_SMALLICON
Return Int($LV_VIEW_SMALLICON)
Case $LV_VIEW_TILE
Return Int($LV_VIEW_TILE)
Case Else
Return -1
EndSwitch
EndFunc
Func _GUICtrlListView_GetViewRect($hWnd)
Local $aRect[4] = [0, 0, 0, 0]
Local $iView = _GUICtrlListView_GetView($hWnd)
If($iView < 0) And($iView > 4) Then Return $aRect
Local $tRECT = DllStructCreate($tagRECT)
__GUICtrl_SendMsg($hWnd, $LVM_GETVIEWRECT, 0, $tRECT, 0, True, -1)
$aRect[0] = DllStructGetData($tRECT, "Left")
$aRect[1] = DllStructGetData($tRECT, "Top")
$aRect[2] = DllStructGetData($tRECT, "Right")
$aRect[3] = DllStructGetData($tRECT, "Bottom")
Return $aRect
EndFunc
Func _GUICtrlListView_HitTest($hWnd, $iX = -1, $iY = -1)
Local $iMode = Opt("MouseCoordMode", 1)
Local $aPos = MouseGetPos()
Opt("MouseCoordMode", $iMode)
Local $tPoint = DllStructCreate($tagPOINT)
DllStructSetData($tPoint, "X", $aPos[0])
DllStructSetData($tPoint, "Y", $aPos[1])
Local $aCall = DllCall("user32.dll", "bool", "ScreenToClient", "hwnd", $hWnd, "struct*", $tPoint)
If @error Then Return SetError(@error, @extended, 0)
If $aCall[0] = 0 Then Return 0
If $iX = -1 Then $iX = DllStructGetData($tPoint, "X")
If $iY = -1 Then $iY = DllStructGetData($tPoint, "Y")
Local $tTest = DllStructCreate($tagLVHITTESTINFO)
DllStructSetData($tTest, "X", $iX)
DllStructSetData($tTest, "Y", $iY)
Local $aTest[10]
$aTest[0] = __GUICtrl_SendMsg($hWnd, $LVM_HITTEST, 0, $tTest, 0, True, -1)
Local $iFlags = DllStructGetData($tTest, "Flags")
$aTest[1] = BitAND($iFlags, $LVHT_NOWHERE) <> 0
$aTest[2] = BitAND($iFlags, $LVHT_ONITEMICON) <> 0
$aTest[3] = BitAND($iFlags, $LVHT_ONITEMLABEL) <> 0
$aTest[4] = BitAND($iFlags, $LVHT_ONITEMSTATEICON) <> 0
$aTest[5] = BitAND($iFlags, $LVHT_ONITEM) <> 0
$aTest[6] = BitAND($iFlags, $LVHT_ABOVE) <> 0
$aTest[7] = BitAND($iFlags, $LVHT_BELOW) <> 0
$aTest[8] = BitAND($iFlags, $LVHT_TOLEFT) <> 0
$aTest[9] = BitAND($iFlags, $LVHT_TORIGHT) <> 0
Return $aTest
EndFunc
Func _GUICtrlListView_InsertColumn($hWnd, $iIndex, $sText, $iWidth = 50, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_INSERTCOLUMNW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_INSERTCOLUMNA
EndIf
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = BitOR($LVCF_FMT, $LVCF_WIDTH, $LVCF_TEXT)
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
Local $iFmt = $aAlign[$iAlign]
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, 1, $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, $tColumn, $tBuffer, False, 4)
If $iAlign > 0 Then _GUICtrlListView_SetColumn($hWnd, $iRet, $sText, $iWidth, $iAlign, $iImage, $bOnRight)
Return $iRet
EndFunc
Func _GUICtrlListView_InsertGroup($hWnd, $iIndex, $iGroupID, $sHeader, $iAlign = 0)
Local $aAlign[3] = [$LVGA_HEADER_LEFT, $LVGA_HEADER_CENTER, $LVGA_HEADER_RIGHT]
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
Local $tHeader = _WinAPI_MultiByteToWideChar($sHeader)
Local $tGroup = DllStructCreate($tagLVGROUP)
Local $iMask = BitOR($LVGF_HEADER, $LVGF_ALIGN, $LVGF_GROUPID)
DllStructSetData($tGroup, "Size", DllStructGetSize($tGroup))
DllStructSetData($tGroup, "Mask", $iMask)
DllStructSetData($tGroup, "GroupID", $iGroupID)
DllStructSetData($tGroup, "Align", $aAlign[$iAlign])
Local $iRet = __GUICtrl_SendMsg($hWnd, $LVM_INSERTGROUP, $iIndex, $tGroup, $tHeader, False, 3)
Return $iRet
EndFunc
Func _GUICtrlListView_InsertItem($hWnd, $sText, $iIndex = -1, $iImage = -1, $iParam = 0)
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_INSERTITEMW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_INSERTITEMA
EndIf
Local $tItem = $__g_tListViewItem
If $iIndex = -1 Then $iIndex = 999999999
DllStructSetData($tBuffer, 1, $sText)
Local $iMask = BitOR($LVIF_TEXT, $LVIF_PARAM)
If $iImage >= 0 Then $iMask = BitOR($iMask, $LVIF_IMAGE)
DllStructSetData($tItem, "Mask", $iMask)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
DllStructSetData($tItem, "Image", $iImage)
DllStructSetData($tItem, "Param", $iParam)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, $tBuffer, False, 6)
Return $iRet
EndFunc
Func _GUICtrlListView_RemoveAllGroups($hWnd)
_GUICtrlListView_BeginUpdate($hWnd)
Local $iGroupID
For $x = _GUICtrlListView_GetGroupCount($hWnd) - 1 To 0 Step -1
$iGroupID = _GUICtrlListView_GetGroupInfoByIndex($hWnd, $x)[2]
_GUICtrlListView_RemoveGroup($hWnd, $iGroupID)
Next
_GUICtrlListView_EndUpdate($hWnd)
EndFunc
Func _GUICtrlListView_RemoveGroup($hWnd, $iGroupID)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_REMOVEGROUP, $iGroupID)
Else
Return GUICtrlSendMsg($hWnd, $LVM_REMOVEGROUP, $iGroupID, 0)
EndIf
EndFunc
Func _GUICtrlListView_Scroll($hWnd, $iDX, $iDY)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_SCROLL, $iDX, $iDY) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_SCROLL, $iDX, $iDY) <> 0
EndIf
EndFunc
Func _GUICtrlListView_SetColumn($hWnd, $iIndex, $sText, $iWidth = -1, $iAlign = -1, $iImage = -1, $bOnRight = False)
Local $aAlign[3] = [$LVCFMT_LEFT, $LVCFMT_RIGHT, $LVCFMT_CENTER]
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_SETCOLUMNW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_SETCOLUMNA
EndIf
Local $tColumn = DllStructCreate($tagLVCOLUMN)
Local $iMask = $LVCF_TEXT
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
$iMask = BitOR($iMask, $LVCF_FMT)
Local $iFmt = $aAlign[$iAlign]
If $iWidth <> -1 Then $iMask = BitOR($iMask, $LVCF_WIDTH)
If $iImage <> -1 Then
$iMask = BitOR($iMask, $LVCF_IMAGE)
$iFmt = BitOR($iFmt, $LVCFMT_COL_HAS_IMAGES, $LVCFMT_IMAGE)
Else
$iImage = 0
EndIf
If $bOnRight Then $iFmt = BitOR($iFmt, $LVCFMT_BITMAP_ON_RIGHT)
DllStructSetData($tBuffer, 1, $sText)
DllStructSetData($tColumn, "Mask", $iMask)
DllStructSetData($tColumn, "Fmt", $iFmt)
DllStructSetData($tColumn, "CX", $iWidth)
DllStructSetData($tColumn, "Image", $iImage)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, $iIndex, $tColumn, $tBuffer, False, 4)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetExtendedListViewStyle($hWnd, $iExStyle, $iExMask = 0)
Local $iRet
If IsHWnd($hWnd) Then
$iRet = _SendMessage($hWnd, $LVM_SETEXTENDEDLISTVIEWSTYLE, $iExMask, $iExStyle)
Else
$iRet = GUICtrlSendMsg($hWnd, $LVM_SETEXTENDEDLISTVIEWSTYLE, $iExMask, $iExStyle)
EndIf
_WinAPI_InvalidateRect($hWnd)
Return $iRet
EndFunc
Func _GUICtrlListView_SetGroupInfo($hWnd, $iGroupID, $sHeader, $iAlign = 0, $iState = $LVGS_NORMAL)
Local $tGroup = 0
If BitAND($iState, $LVGS_SELECTED) Then
$tGroup = __GUICtrlListView_GetGroupInfoEx($hWnd, $iGroupID, BitOR($LVGF_GROUPID, $LVGF_ITEMS))
If @error Or DllStructGetData($tGroup, "cItems") = 0 Then Return False
EndIf
Local $aAlign[3] = [$LVGA_HEADER_LEFT, $LVGA_HEADER_CENTER, $LVGA_HEADER_RIGHT]
If $iAlign < 0 Or $iAlign > 2 Then $iAlign = 0
Local $tHeader = _WinAPI_MultiByteToWideChar($sHeader)
$tGroup = DllStructCreate($tagLVGROUP)
Local $iMask = BitOR($LVGF_HEADER, $LVGF_ALIGN, $LVGF_STATE)
DllStructSetData($tGroup, "Size", DllStructGetSize($tGroup))
DllStructSetData($tGroup, "Mask", $iMask)
DllStructSetData($tGroup, "Align", $aAlign[$iAlign])
DllStructSetData($tGroup, "State", $iState)
DllStructSetData($tGroup, "StateMask", $iState)
Local $iRet = __GUICtrl_SendMsg($hWnd, $LVM_SETGROUPINFO, $iGroupID, $tGroup, $tHeader, False, 3)
DllStructSetData($tGroup, "Mask", $LVGF_GROUPID)
DllStructSetData($tGroup, "GroupID", $iGroupID)
__GUICtrl_SendMsg($hWnd, $LVM_SETGROUPINFO, 0, $tGroup, 0, False, -1)
_WinAPI_InvalidateRect($hWnd)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetItemChecked($hWnd, $iIndex, $bCheck = True)
Local $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$iMsg = $LVM_SETITEMW
Else
$iMsg = $LVM_SETITEMA
EndIf
Local $tItem = $__g_tListViewItem
If($bCheck) Then
DllStructSetData($tItem, "State", 0x2000)
Else
DllStructSetData($tItem, "State", 0x1000)
EndIf
DllStructSetData($tItem, "StateMask", 0xf000)
DllStructSetData($tItem, "Mask", $LVIF_STATE)
DllStructSetData($tItem, "SubItem", 0)
Local $iIndexMax = $iIndex
If $iIndex = -1 Then
$iIndex = 0
$iIndexMax = _GUICtrlListView_GetItemCount($hWnd) - 1
EndIf
Local $iRet
For $x = $iIndex To $iIndexMax
DllStructSetData($tItem, "Item", $x)
$iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, 0, False, -1)
If $iRet = 0 Then ExitLoop
Next
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetItemCount($hWnd, $iItems)
If IsHWnd($hWnd) Then
Return _SendMessage($hWnd, $LVM_SETITEMCOUNT, $iItems, BitOR($LVSICF_NOINVALIDATEALL, $LVSICF_NOSCROLL)) <> 0
Else
Return GUICtrlSendMsg($hWnd, $LVM_SETITEMCOUNT, $iItems, BitOR($LVSICF_NOINVALIDATEALL, $LVSICF_NOSCROLL)) <> 0
EndIf
EndFunc
Func _GUICtrlListView_SetItemEx($hWnd, ByRef $tItem, $iNested = 0)
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_SETITEMW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_SETITEMA
EndIf
Local $iBuffer = 0
If $iNested Then
$tBuffer = 0
DllStructSetData($tItem, "Text", 0)
Else
If DllStructGetData($tItem, "Text") <> -1 Then
$iBuffer = DllStructGetSize($tBuffer)
Else
EndIf
EndIf
DllStructSetData($tItem, "TextMax", $iBuffer)
Local $iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, $tBuffer, False, -1)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetItemGroupID($hWnd, $iIndex, $iGroupID)
Local $tItem = $__g_tListViewItem
DllStructSetData($tItem, "Mask", $LVIF_GROUPID)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
DllStructSetData($tItem, "GroupID", $iGroupID)
Return _GUICtrlListView_SetItemEx($hWnd, $tItem, 1)
EndFunc
Func _GUICtrlListView_SetItemSelected($hWnd, $iIndex, $bSelected = True, $bFocused = False)
Local $tItem = $__g_tListViewItem
Local $iSelected = 0, $iFocused = 0
If($bSelected = True) Then $iSelected = $LVIS_SELECTED
If($bFocused = True And $iIndex <> -1) Then $iFocused = $LVIS_FOCUSED
DllStructSetData($tItem, "Mask", $LVIF_STATE)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", 0)
DllStructSetData($tItem, "State", BitOR($iSelected, $iFocused))
DllStructSetData($tItem, "StateMask", BitOR($LVIS_SELECTED, $iFocused))
Local $iRet = __GUICtrl_SendMsg($hWnd, $LVM_SETITEMSTATE, $iIndex, $tItem, 0, False, -1)
Return $iRet <> 0
EndFunc
Func _GUICtrlListView_SetItemText($hWnd, $iIndex, $sText, $iSubItem = 0)
Local $iRet
If $iSubItem = -1 Then
Local $sSeparatorChar = Opt('GUIDataSeparatorChar')
Local $i_Cols = _GUICtrlListView_GetColumnCount($hWnd)
Local $a_Text = StringSplit($sText, $sSeparatorChar)
If $i_Cols > $a_Text[0] Then $i_Cols = $a_Text[0]
For $i = 1 To $i_Cols
$iRet = _GUICtrlListView_SetItemText($hWnd, $iIndex, $a_Text[$i], $i - 1)
If Not $iRet Then ExitLoop
Next
Return $iRet
EndIf
Local $tBuffer, $iMsg
If _GUICtrlListView_GetUnicodeFormat($hWnd) Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_SETITEMW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_SETITEMA
EndIf
Local $tItem = $__g_tListViewItem
DllStructSetData($tBuffer, 1, $sText)
DllStructSetData($tItem, "Mask", $LVIF_TEXT)
DllStructSetData($tItem, "Item", $iIndex)
DllStructSetData($tItem, "SubItem", $iSubItem)
$iRet = __GUICtrl_SendMsg($hWnd, $iMsg, 0, $tItem, $tBuffer, False, 6, False, -1)
Return $iRet <> 0
EndFunc
#Au3Stripper_Ignore_Funcs=__GUICtrlListView_Sort
Func __GUICtrlListView_Sort($nItem1, $nItem2, $hWnd)
Local $iIndex, $sVal1, $sVal2, $nResult
Local $tBuffer, $iMsg
If $__g_aListViewSortInfo[$iIndex][0] Then
$tBuffer = $__g_tListViewBuffer
$iMsg = $LVM_GETITEMTEXTW
Else
$tBuffer = $__g_tListViewBufferANSI
$iMsg = $LVM_GETITEMTEXTA
EndIf
Local $tItem = $__g_tListViewItem
For $x = 1 To $__g_aListViewSortInfo[0][0]
If $hWnd = $__g_aListViewSortInfo[$x][1] Then
$iIndex = $x
ExitLoop
EndIf
Next
If $__g_aListViewSortInfo[$iIndex][3] = $__g_aListViewSortInfo[$iIndex][4] Then
If Not $__g_aListViewSortInfo[$iIndex][7] Then
$__g_aListViewSortInfo[$iIndex][5] *= -1
$__g_aListViewSortInfo[$iIndex][7] = 1
EndIf
Else
$__g_aListViewSortInfo[$iIndex][7] = 1
EndIf
$__g_aListViewSortInfo[$iIndex][6] = $__g_aListViewSortInfo[$iIndex][3]
DllStructSetData($tItem, "Mask", $LVIF_TEXT)
DllStructSetData($tItem, "SubItem", $__g_aListViewSortInfo[$iIndex][3])
__GUICtrl_SendMsg($hWnd, $iMsg, $nItem1, $tItem, $tBuffer, False, 6, True)
$sVal1 = DllStructGetData($tBuffer, 1)
__GUICtrl_SendMsg($hWnd, $iMsg, $nItem2, $tItem, $tBuffer, False, 6, True)
$sVal2 = DllStructGetData($tBuffer, 1)
If $__g_aListViewSortInfo[$iIndex][8] = 1 Then
If(StringIsFloat($sVal1) Or StringIsInt($sVal1)) Then $sVal1 = Number($sVal1)
If(StringIsFloat($sVal2) Or StringIsInt($sVal2)) Then $sVal2 = Number($sVal2)
EndIf
If $__g_aListViewSortInfo[$iIndex][8] < 2 Then
$nResult = 0
If $sVal1 < $sVal2 Then
$nResult = -1
ElseIf $sVal1 > $sVal2 Then
$nResult = 1
EndIf
Else
$nResult = DllCall('shlwapi.dll', 'int', 'StrCmpLogicalW', 'wstr', $sVal1, 'wstr', $sVal2)[0]
EndIf
$nResult = $nResult * $__g_aListViewSortInfo[$iIndex][5]
Return $nResult
EndFunc
Func _Singleton($sOccurrenceName, $iFlag = 0)
Local Const $ERROR_ALREADY_EXISTS = 183
Local Const $SECURITY_DESCRIPTOR_REVISION = 1
Local $tSecurityAttributes = 0
If BitAND($iFlag, 2) Then
Local $tSecurityDescriptor = DllStructCreate("byte;byte;word;ptr[4]")
Local $aCall = DllCall("advapi32.dll", "bool", "InitializeSecurityDescriptor", "struct*", $tSecurityDescriptor, "dword", $SECURITY_DESCRIPTOR_REVISION)
If @error Then Return SetError(@error, @extended, 0)
If $aCall[0] Then
$aCall = DllCall("advapi32.dll", "bool", "SetSecurityDescriptorDacl", "struct*", $tSecurityDescriptor, "bool", 1, "ptr", 0, "bool", 0)
If @error Then Return SetError(@error, @extended, 0)
If $aCall[0] Then
$tSecurityAttributes = DllStructCreate($tagSECURITY_ATTRIBUTES)
DllStructSetData($tSecurityAttributes, 1, DllStructGetSize($tSecurityAttributes))
DllStructSetData($tSecurityAttributes, 2, DllStructGetPtr($tSecurityDescriptor))
DllStructSetData($tSecurityAttributes, 3, 0)
EndIf
EndIf
EndIf
Local $aHandle = DllCall("kernel32.dll", "handle", "CreateMutexW", "struct*", $tSecurityAttributes, "bool", 1, "wstr", $sOccurrenceName)
If @error Then Return SetError(@error, @extended, 0)
Local $aLastError = DllCall("kernel32.dll", "dword", "GetLastError")
If @error Then Return SetError(@error, @extended, 0)
If $aLastError[0] = $ERROR_ALREADY_EXISTS Then
If BitAND($iFlag, 1) Then
DllCall("kernel32.dll", "bool", "CloseHandle", "handle", $aHandle[0])
If @error Then Return SetError(@error, @extended, 0)
Return SetError($aLastError[0], $aLastError[0], 0)
Else
Exit -1
EndIf
EndIf
Return $aHandle[0]
EndFunc
AutoItSetOption("GUICloseOnESC", 0)
Global Const $g_AppWndTitle = "AdobeGenP", $g_AppVersion = "Original version by uncia - CGP Community Edition - v3.2.0"
If _Singleton($g_AppWndTitle, 1) = 0 Then
Exit
EndIf
Global $MyLVGroupIsExpanded = True
Global $fInterrupt = 0
Global $FilesToPatch[0][1], $FilesToPatchNull[0][1]
Global $FilesToRestore[0][1], $fFilesListed = 0
Global $MyhGUI, $idMsg, $idListview, $g_idListview, $idButtonSearch, $idButtonStop
Global $idButtonCustomFolder, $idBtnCure, $idBtnDeselectAll, $ListViewSelectFlag = 1
Global $idBtnBlockPopUp, $idBtnRunasTI, $idMemo, $timestamp, $idLog, $idBtnRestore
Global $sINIPath = @ScriptDir & "\config.ini"
If Not FileExists($sINIPath) Then
FileInstall("config.ini", @ScriptDir & "\config.ini")
EndIf
Global $MyDefPath = IniRead($sINIPath, "Default", "Path", "C:\Program Files")
If Not FileExists($MyDefPath) Or Not StringInStr(FileGetAttrib($MyDefPath), "D") Then
IniWrite($sINIPath, "Default", "Path", "C:\Program Files")
$MyDefPath = "C:\Program Files"
EndIf
If(@UserName = "SYSTEM") Then
FileDelete(@WindowsDir & "\Temp\RunAsTI.exe")
EndIf
Global $MyRegExpGlobalPatternSearchCount = 0, $Count = 0, $idProgressBar
Global $aOutHexGlobalArray[0], $aNullArray[0], $aInHexArray[0]
Global $sz_type, $bFoundAcro32 = False, $aSpecialFiles, $sSpecialFiles = "|"
Global $ProgressFileCountScale, $FileSearchedCount
Local $tTargetFileList_Adobe = IniReadSection($sINIPath, "TargetFiles")
Global $TargetFileList_Adobe[0]
If Not @error Then
ReDim $TargetFileList_Adobe[$tTargetFileList_Adobe[0][0]]
For $i = 1 To $tTargetFileList_Adobe[0][0]
$TargetFileList_Adobe[$i - 1] = StringReplace($tTargetFileList_Adobe[$i][1], '"', "")
Next
EndIf
$aSpecialFiles = IniReadSection($sINIPath, "CustomPatterns")
For $i = 1 To UBound($aSpecialFiles) - 1
$sSpecialFiles = $sSpecialFiles & $aSpecialFiles[$i][0] & "|"
Next
GUIRegisterMsg($WM_COMMAND, "WM_COMMAND")
MainGui()
While 1
$idMsg = GUIGetMsg()
Select
Case $idMsg = $GUI_EVENT_CLOSE
GUIDelete($MyhGUI)
_Exit()
Case $idMsg = $GUI_EVENT_RESIZED
ContinueCase
Case $idMsg = $GUI_EVENT_RESTORE
ContinueCase
Case $idMsg = $GUI_EVENT_MAXIMIZE
Local $iWidth
Local $aGui = WinGetPos($MyhGUI)
Local $aRect = _GUICtrlListView_GetViewRect($g_idListview)
If($aRect[2] > $aGui[2]) Then
$iWidth = $aGui[2] - 75
Else
$iWidth = $aRect[2] - 25
EndIf
GUICtrlSendMsg($idListview, $LVM_SETCOLUMNWIDTH, 1, $iWidth)
Case $idMsg = $idButtonStop
$ListViewSelectFlag = 0
FillListViewWithInfo()
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
GUICtrlSetState($idButtonStop, $GUI_HIDE)
GUICtrlSetState($idButtonSearch, $GUI_SHOW)
GUICtrlSetState($idButtonSearch, 64)
GUICtrlSetState($idBtnRestore, $GUI_HIDE)
GUICtrlSetState($idBtnBlockPopUp, $GUI_SHOW)
GUICtrlSetState($idBtnBlockPopUp, 64)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idBtnCure, 128)
Case $idMsg = $idButtonSearch
$fInterrupt = 0
GUICtrlSetState($idButtonSearch, $GUI_HIDE)
GUICtrlSetState($idButtonStop, $GUI_SHOW)
ToggleLog(0)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idBtnBlockPopUp, 128)
GUICtrlSetState($idListview, 128)
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idButtonCustomFolder, 128)
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
_GUICtrlListView_AddItem($idListview, "", 0)
_GUICtrlListView_AddItem($idListview, "", 1)
_GUICtrlListView_AddItem($idListview, "", 2)
_GUICtrlListView_AddItem($idListview, "", 2)
_GUICtrlListView_RemoveAllGroups($idListview)
_GUICtrlListView_InsertGroup($idListview, -1, 1, "", 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
_GUICtrlListView_AddSubItem($idListview, 0, "", 1)
_GUICtrlListView_AddSubItem($idListview, 1, "Preparing...", 1)
_GUICtrlListView_AddSubItem($idListview, 2, "", 1)
_GUICtrlListView_AddSubItem($idListview, 3, "Be patient, please.", 1)
_GUICtrlListView_SetItemGroupID($idListview, 0, 1)
_GUICtrlListView_SetItemGroupID($idListview, 1, 1)
_GUICtrlListView_SetItemGroupID($idListview, 2, 1)
_GUICtrlListView_SetItemGroupID($idListview, 3, 1)
_Expand_All_Click()
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
$FilesToPatch = $FilesToPatchNull
$FilesToRestore = $FilesToPatchNull
$timestamp = TimerInit()
Local $FileCount
Local $aSize = DirGetSize($MyDefPath, $DIR_EXTENDED)
If UBound($aSize) >= 2 Then
$FileCount = $aSize[1]
$ProgressFileCountScale = 100 / $FileCount
$FileSearchedCount = 0
ProgressWrite(0)
RecursiveFileSearch($MyDefPath, 0, $FileCount)
Sleep(100)
ProgressWrite(0)
EndIf
If $MyDefPath = "C:\Program Files" Or $MyDefPath = "C:\Program Files\Adobe" Then
Local $sAppsPanelDir = EnvGet('ProgramFiles(x86)') & "\Common Files\Adobe"
$aSize = DirGetSize($sAppsPanelDir, $DIR_EXTENDED)
If UBound($aSize) >= 2 Then
$FileCount = $aSize[1]
RecursiveFileSearch($sAppsPanelDir, 0, $FileCount)
ProgressWrite(0)
EndIf
EndIf
FillListViewWithFiles()
If _GUICtrlListView_GetItemCount($idListview) > 0 Then
_Assign_Groups_To_Found_Files()
$ListViewSelectFlag = 1
GUICtrlSetState($idButtonSearch, 128)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idBtnCure, 64)
GUICtrlSetState($idBtnCure, 256)
If UBound($FilesToRestore) > 0 Then
GUICtrlSetState($idBtnBlockPopUp, $GUI_HIDE)
GUICtrlSetState($idBtnRestore, 64)
GUICtrlSetState($idBtnRestore, $GUI_SHOW)
EndIf
Else
$ListViewSelectFlag = 0
FillListViewWithInfo()
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idButtonSearch, 64)
GUICtrlSetState($idButtonSearch, 256)
EndIf
_Expand_All_Click()
GUICtrlSetState($idBtnDeselectAll, 64)
GUICtrlSetState($idBtnBlockPopUp, 64)
GUICtrlSetState($idListview, 64)
GUICtrlSetState($idButtonCustomFolder, 64)
GUICtrlSetState($idButtonSearch, $GUI_SHOW)
GUICtrlSetState($idButtonStop, $GUI_HIDE)
Case $idMsg = $idButtonCustomFolder
ToggleLog(0)
MyFileOpenDialog()
_Expand_All_Click()
If $fFilesListed = 0 Then
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idButtonSearch, 64)
GUICtrlSetState($idButtonSearch, 256)
Else
GUICtrlSetState($idButtonSearch, 128)
GUICtrlSetState($idBtnDeselectAll, 64)
GUICtrlSetState($idBtnCure, 64)
GUICtrlSetState($idBtnCure, 256)
EndIf
Case $idMsg = $idBtnDeselectAll
ToggleLog(0)
If $ListViewSelectFlag = 1 Then
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
_GUICtrlListView_SetItemChecked($idListview, $i, 0)
Next
$ListViewSelectFlag = 0
Else
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
_GUICtrlListView_SetItemChecked($idListview, $i, 1)
Next
$ListViewSelectFlag = 1
EndIf
Case $idMsg = $idBtnBlockPopUp
ToggleLog(0)
BlockPopUp()
Case $idMsg = $idBtnRunasTI
FileInstall("RunAsTI.exe", @WindowsDir & "\Temp\RunAsTI.exe")
Exit Run(@WindowsDir & '\Temp\RunAsTI.exe "' & @ScriptFullPath & '"')
Case $idMsg = $idBtnCure
ToggleLog(0)
GUICtrlSetState($idListview, 128)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idButtonSearch, 128)
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idBtnBlockPopUp, 128)
GUICtrlSetState($idBtnRestore, 128)
GUICtrlSetState($idButtonCustomFolder, 128)
_Expand_All_Click()
_GUICtrlListView_EnsureVisible($idListview, 0, 0)
Local $ItemFromList
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
If _GUICtrlListView_GetItemChecked($idListview, $i) = True Then
_GUICtrlListView_SetItemSelected($idListview, $i)
$ItemFromList = _GUICtrlListView_GetItemText($idListview, $i, 1)
MyGlobalPatternSearch($ItemFromList)
ProgressWrite(0)
Sleep(100)
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $ItemFromList & @CRLF & "---" & @CRLF & "medication :)")
LogWrite(1, $ItemFromList)
Sleep(100)
MyGlobalPatternPatch($ItemFromList, $aOutHexGlobalArray)
_GUICtrlListView_Scroll($idListview, 0, 10)
_GUICtrlListView_EnsureVisible($idListview, $i, 0)
Sleep(100)
EndIf
_GUICtrlListView_SetItemChecked($idListview, $i, False)
Next
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
_GUICtrlListView_RemoveAllGroups($idListview)
_GUICtrlListView_InsertGroup($idListview, -1, 1, "", 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
GUICtrlSetState($idListview, 64)
GUICtrlSetState($idButtonSearch, 64)
GUICtrlSetState($idButtonCustomFolder, 64)
GUICtrlSetState($idBtnBlockPopUp, 64)
GUICtrlSetState($idBtnBlockPopUp, $GUI_SHOW)
GUICtrlSetState($idBtnRestore, $GUI_HIDE)
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idButtonSearch, 256)
FillListViewWithInfo()
If $bFoundAcro32 = True Then
MsgBox($MB_SYSTEMMODAL, "Information", "GenP does not patch the x32 bit version of Acrobat. Please use the x64 bit version of Acrobat.")
LogWrite(1, "GenP does not patch the x32 bit version of Acrobat. Please use the x64 bit version of Acrobat.")
EndIf
ToggleLog(1)
Case $idMsg = $idBtnRestore
GUICtrlSetData($idLog, "Activity Log")
ToggleLog(0)
GUICtrlSetState($idListview, 128)
GUICtrlSetState($idBtnDeselectAll, 128)
GUICtrlSetState($idButtonSearch, 128)
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idBtnBlockPopUp, 128)
GUICtrlSetState($idBtnRestore, 128)
GUICtrlSetState($idButtonCustomFolder, 128)
_Expand_All_Click()
_GUICtrlListView_EnsureVisible($idListview, 0, 0)
Local $ItemFromList, $iCheckedItems, $iProgress
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
If _GUICtrlListView_GetItemChecked($idListview, $i) = True Then
_GUICtrlListView_SetItemSelected($idListview, $i)
$ItemFromList = _GUICtrlListView_GetItemText($idListview, $i, 1)
$iCheckedItems = _GUICtrlListView_GetSelectedCount($idListview)
$iProgress = 100 / $iCheckedItems
ProgressWrite(0)
RestoreFile($ItemFromList)
ProgressWrite($iProgress)
Sleep(100)
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $ItemFromList & @CRLF & "---" & @CRLF & "restoring :)")
Sleep(100)
_GUICtrlListView_Scroll($idListview, 0, 10)
_GUICtrlListView_EnsureVisible($idListview, $i, 0)
Sleep(100)
EndIf
_GUICtrlListView_SetItemChecked($idListview, $i, False)
Next
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
_GUICtrlListView_RemoveAllGroups($idListview)
_GUICtrlListView_InsertGroup($idListview, -1, 1, "", 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
GUICtrlSetState($idListview, 64)
GUICtrlSetState($idButtonCustomFolder, 64)
GUICtrlSetState($idBtnBlockPopUp, 64)
GUICtrlSetState($idBtnBlockPopUp, $GUI_SHOW)
GUICtrlSetState($idBtnRestore, $GUI_HIDE)
GUICtrlSetState($idBtnRestore, 64)
GUICtrlSetState($idBtnCure, 128)
GUICtrlSetState($idButtonSearch, 64)
GUICtrlSetState($idButtonSearch, 256)
FillListViewWithInfo()
ToggleLog(1)
EndSelect
WEnd
Func MainGui()
$MyhGUI = GUICreate($g_AppWndTitle, 595, 805, -1, -1, BitOR($WS_MAXIMIZEBOX, $WS_MINIMIZEBOX, $WS_SIZEBOX, $GUI_SS_DEFAULT_GUI))
$idListview = GUICtrlCreateListView("", 10, 10, 575, 580)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$g_idListview = GUICtrlGetHandle($idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER, $LVS_EX_CHECKBOXES))
_GUICtrlListView_SetItemCount($idListview, UBound($FilesToPatch))
_GUICtrlListView_AddColumn($idListview, "", 20)
_GUICtrlListView_AddColumn($idListview, "For collapsing or expanding all groups, please click here", 532, 2)
_GUICtrlListView_EnableGroupView($idListview)
_GUICtrlListView_InsertGroup($idListview, -1, 1, "", 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
FillListViewWithInfo()
$idButtonCustomFolder = GUICtrlCreateButton("Path", 10, 630, 80, 30)
GUICtrlSetTip(-1, "Select Path that You want -> press Search -> press Patch button")
GUICtrlSetImage(-1, "imageres.dll", -4, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idButtonSearch = GUICtrlCreateButton("Search", 110, 630, 80, 30)
GUICtrlSetTip(-1, "Let GenP find Apps automatically in current path")
GUICtrlSetImage(-1, "imageres.dll", -8, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idButtonStop = GUICtrlCreateButton("Stop", 110, 630, 80, 30)
GUICtrlSetState(-1, $GUI_HIDE)
GUICtrlSetTip(-1, "Stop searching for Apps")
GUICtrlSetImage(-1, "imageres.dll", -8, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idBtnDeselectAll = GUICtrlCreateButton("De/Select", 210, 630, 80, 30)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetTip(-1, "De/Select All files")
GUICtrlSetImage(-1, "imageres.dll", -76, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idBtnCure = GUICtrlCreateButton("Patch", 305, 630, 80, 30)
GUICtrlSetState(-1, $GUI_DISABLE)
GUICtrlSetTip(-1, "Patch all selected files")
GUICtrlSetImage(-1, "imageres.dll", -102, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idBtnBlockPopUp = GUICtrlCreateButton("Pop-up", 405, 630, 80, 30)
GUICtrlSetTip(-1, "Block Unlicensed Pop-up by creating Windows Firewall Rule")
GUICtrlSetImage(-1, "imageres.dll", -101, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idBtnRestore = GUICtrlCreateButton("Restore", 405, 630, 80, 30)
GUICtrlSetState(-1, $GUI_HIDE)
GUICtrlSetTip(-1, "Restore Original Files")
GUICtrlSetImage(-1, "imageres.dll", -113, 0)
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idBtnRunasTI = GUICtrlCreateButton("Runas TI", 505, 630, 80, 30)
GUICtrlSetImage(-1, "imageres.dll", -74, 0)
If(@UserName = "SYSTEM") Then
GUICtrlSetState(-1, $GUI_DISABLE)
EndIf
GUICtrlSetTip(-1, "Run as Trusted Installer")
GUICtrlSetResizing(-1, $GUI_DOCKAUTO)
$idProgressBar = GUICtrlCreateProgress(10, 597, 575, 25, $PBS_SMOOTHREVERSE)
GUICtrlSetResizing(-1, $GUI_DOCKVCENTER)
$idMemo = GUICtrlCreateEdit("", 10, 670, 575, 100, BitOR($ES_READONLY, $ES_CENTER, $WS_DISABLED))
GUICtrlSetResizing(-1, $GUI_DOCKVCENTER)
$idLog = GUICtrlCreateEdit("", 10, 670, 575, 100, BitOR($WS_VSCROLL, $ES_AUTOVSCROLL, $ES_READONLY))
GUICtrlSetResizing(-1, $GUI_DOCKVCENTER)
GUICtrlSetState($idLog, $GUI_HIDE)
GUICtrlSetData($idLog, "Activity Log")
GUICtrlCreateLabel($g_AppVersion, 10, 780, 575, 25, $ES_CENTER)
GUICtrlSetResizing(-1, $GUI_DOCKBOTTOM)
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
GUICtrlSetState($idButtonSearch, 256)
GUISetState(@SW_SHOW)
GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
EndFunc
Func RecursiveFileSearch($INSTARTDIR, $DEPTH, $FileCount)
_GUICtrlListView_SetItemText($idListview, 1, "Searching for files.", 1)
Local $RecursiveFileSearch_MaxDeep = 6
If $DEPTH > $RecursiveFileSearch_MaxDeep Then Return
Local $STARTDIR = $INSTARTDIR & "\"
$FileSearchedCount += 1
Local $HSEARCH = FileFindFirstFile($STARTDIR & "*.*")
If @error Then Return
Local $NEXT, $IPATH, $isDir
While $fInterrupt = 0
$NEXT = FileFindNextFile($HSEARCH)
$FileSearchedCount += 1
If @error Then ExitLoop
$isDir = StringInStr(FileGetAttrib($STARTDIR & $NEXT), "D")
If $isDir Then
Local $targetDepth
$targetDepth = RecursiveFileSearch($STARTDIR & $NEXT, $DEPTH + 1, $FileCount)
Else
$IPATH = $STARTDIR & $NEXT
Local $FileNameCropped
If(IsArray($TargetFileList_Adobe)) Then
For $AdobeFileTarget In $TargetFileList_Adobe
$FileNameCropped = StringSplit(StringLower($IPATH), StringLower($AdobeFileTarget), $STR_ENTIRESPLIT)
If @error <> 1 Then
If Not StringInStr($IPATH, ".bak") Then
If StringInStr($IPATH, "Adobe") Or StringInStr($IPATH, "Acrobat") Then
_ArrayAdd($FilesToPatch, $IPATH)
EndIf
Else
_ArrayAdd($FilesToRestore, $IPATH)
EndIf
EndIf
Next
EndIf
EndIf
WEnd
If 1 = Random(0, 10, 1) Then
MemoWrite(@CRLF & "Searching in " & $FileCount & " files" & @TAB & @TAB & "Found : " & UBound($FilesToPatch) & @CRLF & "---" & @CRLF & "Level: " & $DEPTH & " Time elapsed : " & Round(TimerDiff($timestamp) / 1000, 0) & " second(s)" & @TAB & @TAB & "Excluded because of *.bak: " & UBound($FilesToRestore) & @CRLF & "---" & @CRLF & $INSTARTDIR )
ProgressWrite($ProgressFileCountScale * $FileSearchedCount)
EndIf
FileClose($HSEARCH)
EndFunc
Func FillListViewWithInfo()
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER))
_Expand_All_Click()
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
For $i = 0 To 24
_GUICtrlListView_AddItem($idListview, "", $i)
_GUICtrlListView_SetItemGroupID($idListview, $i, 1)
Next
_GUICtrlListView_AddSubItem($idListview, 0, "", 1)
_GUICtrlListView_AddSubItem($idListview, 1, "To patch all Adobe apps in default location:", 1)
_GUICtrlListView_AddSubItem($idListview, 2, "Press 'Search Files' - Press 'Patch Files'", 1)
_GUICtrlListView_AddSubItem($idListview, 3, "Default path - C:\Program Files", 1)
_GUICtrlListView_AddSubItem($idListview, 4, '-------------------------------------------------------------', 1)
_GUICtrlListView_AddSubItem($idListview, 5, "After searching, some products may already be patched.", 1)
_GUICtrlListView_AddSubItem($idListview, 6, "To select\deselect products to patch, LEFT CLICK on the product group", 1)
_GUICtrlListView_AddSubItem($idListview, 7, "To select\deselect individual files, RIGHT CLICK on the file", 1)
_GUICtrlListView_AddSubItem($idListview, 8, '-------------------------------------------------------------', 1)
_GUICtrlListView_AddSubItem($idListview, 9, "What's new in GenP:", 1)
_GUICtrlListView_AddSubItem($idListview, 10, "Can patch apps from 2019 version to current and future releases", 1)
_GUICtrlListView_AddSubItem($idListview, 11, "Automatic search and patch in selected folder", 1)
_GUICtrlListView_AddSubItem($idListview, 12, "New patching logic. 'Unlicensed Pop-up' Blocker for Windows Firewall", 1)
_GUICtrlListView_AddSubItem($idListview, 13, "Support for all Substance products", 1)
_GUICtrlListView_AddSubItem($idListview, 14, '-------------------------------------------------------------', 1)
_GUICtrlListView_AddSubItem($idListview, 15, "Known issues:", 1)
_GUICtrlListView_AddSubItem($idListview, 16, "InDesign and InCopy will have high Cpu usage", 1)
_GUICtrlListView_AddSubItem($idListview, 17, "Animate will have some problems with home screen if Signed Out", 1)
_GUICtrlListView_AddSubItem($idListview, 18, "Acrobat, XD, Lightroom Classic will partially work if Signed Out", 1)
_GUICtrlListView_AddSubItem($idListview, 19, "Premiere Rush, Lightroom Online, Photoshop Express", 1)
_GUICtrlListView_AddSubItem($idListview, 20, "Won't be fully unlocked", 1)
_GUICtrlListView_AddSubItem($idListview, 21, '-------------------------------------------------------------', 1)
_GUICtrlListView_AddSubItem($idListview, 22, "Some Apps demand Creative Cloud App and mandatory SignIn", 1)
_GUICtrlListView_AddSubItem($idListview, 23, "Fresco, Aero, Lightroom Online, Premiere Rush, Photoshop Express", 1)
$fFilesListed = 0
EndFunc
Func FillListViewWithFiles()
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_FULLROWSELECT, $LVS_EX_GRIDLINES, $LVS_EX_DOUBLEBUFFER, $LVS_EX_CHECKBOXES))
If UBound($FilesToPatch) > 0 Then
Global $aItems[UBound($FilesToPatch)][2]
For $i = 0 To UBound($aItems) - 1
$aItems[$i][0] = $i
$aItems[$i][1] = $FilesToPatch[$i][0]
Next
_GUICtrlListView_AddArray($idListview, $aItems)
MemoWrite(@CRLF & UBound($FilesToPatch) & " File(s) were found in " & Round(TimerDiff($timestamp) / 1000, 0) & " second(s) at:" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "Press the 'Patch Files'")
LogWrite(1, UBound($FilesToPatch) & " File(s) were found in " & Round(TimerDiff($timestamp) / 1000, 0) & " second(s)" & @CRLF)
$fFilesListed = 1
Else
MemoWrite(@CRLF & "Nothing was found in" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
LogWrite(1, "Nothing was found in " & $MyDefPath)
$fFilesListed = 0
EndIf
EndFunc
Func MemoWrite($sMessage)
GUICtrlSetData($idMemo, $sMessage)
EndFunc
Func LogWrite($bTS, $sMessage)
GUICtrlSetDataEx($idLog, $sMessage, $bTS)
EndFunc
Func ToggleLog($bShow)
If $bShow = 1 Then
GUICtrlSetState($idMemo, $GUI_HIDE)
GUICtrlSetState($idLog, $GUI_SHOW)
Else
GUICtrlSetState($idLog, $GUI_HIDE)
GUICtrlSetState($idMemo, $GUI_SHOW)
EndIf
EndFunc
Func GUICtrlSetDataEx($hWnd, $sText, $bTS)
If Not IsHWnd($hWnd) Then $hWnd = GUICtrlGetHandle($hWnd)
Local $iLength = DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0x000E, "wparam", 0, "lparam", 0)
DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0xB1, "wparam", $iLength[0], "lparam", $iLength[0])
If $bTS = 1 Then
Local $iData = @CRLF & @YEAR & "-" & @MON & "-" & @MDAY & " " & @HOUR & ":" & @MIN & ":" & @SEC & "." & @MSEC & " " & $sText
Else
Local $iData = $sText
EndIf
DllCall("user32.dll", "lresult", "SendMessageW", "hwnd", $hWnd, "uint", 0xC2, "wparam", True, "wstr", $iData)
EndFunc
Func ProgressWrite($msg_Progress)
GUICtrlSetData($idProgressBar, $msg_Progress)
EndFunc
Func MyFileOpenDialog()
Local Const $sMessage = "Select a Path"
FileSetAttrib("C:\Program Files\WindowsApps", "-H")
Local $MyTempPath = FileSelectFolder($sMessage, $MyDefPath, 0, $MyDefPath, $MyhGUI)
If @error Then
FileSetAttrib("C:\Program Files\WindowsApps", "+H")
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "waiting for user action")
Else
GUICtrlSetState($idBtnCure, 128)
$MyDefPath = $MyTempPath
IniWrite($sINIPath, "Default", "Path", $MyDefPath)
_GUICtrlListView_DeleteAllItems($g_idListview)
_GUICtrlListView_SetExtendedListViewStyle($idListview, BitOR($LVS_EX_GRIDLINES, $LVS_EX_FULLROWSELECT, $LVS_EX_SUBITEMIMAGES))
_GUICtrlListView_AddItem($idListview, "", 0)
_GUICtrlListView_AddItem($idListview, "", 1)
_GUICtrlListView_AddItem($idListview, "", 2)
_GUICtrlListView_AddItem($idListview, "", 3)
_GUICtrlListView_AddItem($idListview, "", 4)
_GUICtrlListView_AddItem($idListview, "", 5)
_GUICtrlListView_AddItem($idListview, "", 6)
_GUICtrlListView_AddSubItem($idListview, 0, "", 1)
_GUICtrlListView_AddSubItem($idListview, 1, "Path:", 1)
_GUICtrlListView_AddSubItem($idListview, 2, " " & $MyDefPath, 1)
_GUICtrlListView_AddSubItem($idListview, 3, "Step 1:", 1)
_GUICtrlListView_AddSubItem($idListview, 4, " Press 'Search Files' - wait until GenP finds all files", 1)
_GUICtrlListView_AddSubItem($idListview, 5, "Step 2:", 1)
_GUICtrlListView_AddSubItem($idListview, 6, " Press 'Patch Files' - wait until GenP will do it's job", 1)
_GUICtrlListView_SetItemGroupID($idListview, 0, 1)
_GUICtrlListView_SetItemGroupID($idListview, 1, 1)
_GUICtrlListView_SetItemGroupID($idListview, 2, 1)
_GUICtrlListView_SetItemGroupID($idListview, 3, 1)
_GUICtrlListView_SetItemGroupID($idListview, 4, 1)
_GUICtrlListView_SetItemGroupID($idListview, 5, 1)
_GUICtrlListView_SetItemGroupID($idListview, 6, 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Info", 1, $LVGS_COLLAPSIBLE)
FileSetAttrib("C:\Program Files\WindowsApps", "+H")
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyDefPath & @CRLF & "---" & @CRLF & "Press the Search button")
GUICtrlSetState($idBtnBlockPopUp, $GUI_SHOW)
GUICtrlSetState($idBtnRestore, $GUI_HIDE)
$fFilesListed = 0
EndIf
EndFunc
Func _ProcessCloseEx($sName)
Local $iPID = Run("TASKKILL /F /T /IM " & $sName, @TempDir, @SW_HIDE)
ProcessWaitClose($iPID)
EndFunc
Func MyGlobalPatternSearch($MyFileToParse)
$aInHexArray = $aNullArray
$aOutHexGlobalArray = $aNullArray
ProgressWrite(0)
$MyRegExpGlobalPatternSearchCount = 0
$Count = 15
Local $sFileName = StringRegExpReplace($MyFileToParse, "^.*\\", "")
Local $sExt = StringRegExpReplace($sFileName, "^.*\.", "")
MemoWrite(@CRLF & $MyFileToParse & @CRLF & "---" & @CRLF & "Preparing to Analyze" & @CRLF & "---" & @CRLF & "*****")
LogWrite(1, "Checking File: " & $sFileName & " ")
If $sExt = "exe" Then
_ProcessCloseEx("""" & $sFileName & """")
EndIf
If $sFileName = "AppsPanelBL.dll" Then
_ProcessCloseEx("""Creative Cloud.exe""")
_ProcessCloseEx("""Adobe Desktop Service.exe""")
Sleep(100)
EndIf
If StringInStr($sSpecialFiles, $sFileName) Then
LogWrite(0, " - using Custom Patterns")
ExecuteSearchPatterns($sFileName, 0, $MyFileToParse)
Else
LogWrite(0, " - using Default Patterns")
ExecuteSearchPatterns($sFileName, 1, $MyFileToParse)
EndIf
Sleep(100)
EndFunc
Func ExecuteSearchPatterns($FileName, $DefaultPatterns, $MyFileToParse)
Local $aPatterns, $sPattern, $sData, $aArray, $sSearch, $sReplace, $iPatternLength
If $DefaultPatterns = 0 Then
$aPatterns = IniReadArray($sINIPath, "CustomPatterns", $FileName, "")
Else
$aPatterns = IniReadArray($sINIPath, "DefaultPatterns", "Values", "")
EndIf
For $i = 0 To UBound($aPatterns) - 1
$sPattern = $aPatterns[$i]
$sData = IniRead($sINIPath, "Patches", $sPattern, "")
If StringInStr($sData, "|") Then
$aArray = StringSplit($sData, "|")
If UBound($aArray) = 3 Then
$sSearch = StringReplace($aArray[1], '"', '')
$sReplace = StringReplace($aArray[2], '"', '')
$iPatternLength = StringLen($sSearch)
If $iPatternLength <> StringLen($sReplace) Or Mod($iPatternLength, 2) <> 0 Then
MsgBox($MB_SYSTEMMODAL, "Error", "Pattern Error in config.ini:" & $sPattern & @CRLF & $sSearch & @CRLF & $sReplace)
Exit
EndIf
LogWrite(1, "Searching for: " & $sPattern & ": " & $sSearch)
MyRegExpGlobalPatternSearch($MyFileToParse, $sSearch, $sReplace, $sPattern)
EndIf
EndIf
Next
EndFunc
Func MyRegExpGlobalPatternSearch($FileToParse, $PatternToSearch, $PatternToReplace, $PatternName)
Local $hFileOpen = FileOpen($FileToParse, $FO_READ + $FO_BINARY)
FileSetPos($hFileOpen, 60, 0)
$sz_type = FileRead($hFileOpen, 4)
FileSetPos($hFileOpen, Number($sz_type) + 4, 0)
$sz_type = FileRead($hFileOpen, 2)
If $sz_type = "0x4C01" And StringInStr($FileToParse, "Acrobat", 2) > 0 Then
MemoWrite(@CRLF & $FileToParse & @CRLF & "---" & @CRLF & "File is 32bit. Aborting..." & @CRLF & "---")
FileClose($hFileOpen)
Sleep(100)
$bFoundAcro32 = True
Else
FileSetPos($hFileOpen, 0, 0)
Local $sFileRead = FileRead($hFileOpen)
Local $GeneQuestionMark, $AnyNumOfBytes, $OutStringForRegExp
For $i = 256 To 1 Step -2
$GeneQuestionMark = _StringRepeat("??", $i / 2)
$AnyNumOfBytes = "(.{" & $i & "})"
$OutStringForRegExp = StringReplace($PatternToSearch, $GeneQuestionMark, $AnyNumOfBytes)
$PatternToSearch = $OutStringForRegExp
Next
Local $sSearchPattern = $OutStringForRegExp
Local $aReplacePattern = $PatternToReplace
Local $sWildcardSearchPattern = "", $sWildcardReplacePattern = "", $sFinalReplacePattern = ""
Local $aInHexTempArray[0]
Local $sSearchCharacter = "", $sReplaceCharacter = ""
$aInHexTempArray = $aNullArray
$aInHexTempArray = StringRegExp($sFileRead, $sSearchPattern, $STR_REGEXPARRAYGLOBALFULLMATCH, 1)
For $i = 0 To UBound($aInHexTempArray) - 1
$aInHexArray = $aNullArray
$sSearchCharacter = ""
$sReplaceCharacter = ""
$sWildcardSearchPattern = ""
$sWildcardReplacePattern = ""
$sFinalReplacePattern = ""
$aInHexArray = $aInHexTempArray[$i]
If @error = 0 Then
$sWildcardSearchPattern = $aInHexArray[0]
$sWildcardReplacePattern = $aReplacePattern
If StringInStr($sWildcardReplacePattern, "?") Then
For $j = 1 To StringLen($sWildcardReplacePattern) + 1
$sSearchCharacter = StringMid($sWildcardSearchPattern, $j, 1)
$sReplaceCharacter = StringMid($sWildcardReplacePattern, $j, 1)
If $sReplaceCharacter <> "?" Then
$sFinalReplacePattern &= $sReplaceCharacter
Else
$sFinalReplacePattern &= $sSearchCharacter
EndIf
Next
Else
$sFinalReplacePattern = $sWildcardReplacePattern
EndIf
_ArrayAdd($aOutHexGlobalArray, $sWildcardSearchPattern)
_ArrayAdd($aOutHexGlobalArray, $sFinalReplacePattern)
ConsoleWrite($PatternName & "---" & @TAB & $sWildcardSearchPattern & "	" & @CRLF)
ConsoleWrite($PatternName & "R" & "--" & @TAB & $sFinalReplacePattern & "	" & @CRLF)
MemoWrite(@CRLF & $FileToParse & @CRLF & "---" & @CRLF & $PatternName & @CRLF & "---" & @CRLF & $sWildcardSearchPattern & @CRLF & $sFinalReplacePattern)
LogWrite(1, "Replacing with: " & $sFinalReplacePattern)
Else
ConsoleWrite($PatternName & "---" & @TAB & "No" & "	" & @CRLF)
MemoWrite(@CRLF & $FileToParse & @CRLF & "---" & @CRLF & $PatternName & "---" & "No")
EndIf
$MyRegExpGlobalPatternSearchCount += 1
Next
FileClose($hFileOpen)
$sFileRead = ""
ProgressWrite(Round($MyRegExpGlobalPatternSearchCount / $Count * 100))
Sleep(100)
EndIf
EndFunc
Func MyGlobalPatternPatch($MyFileToPatch, $MyArrayToPatch)
ProgressWrite(0)
Local $iRows = UBound($MyArrayToPatch)
If $iRows > 0 Then
MemoWrite(@CRLF & "Path" & @CRLF & "---" & @CRLF & $MyFileToPatch & @CRLF & "---" & @CRLF & "medication :)")
Local $hFileOpen = FileOpen($MyFileToPatch, $FO_READ + $FO_BINARY)
Local $sFileRead = FileRead($hFileOpen)
Local $sStringOut
For $i = 0 To $iRows - 1 Step 2
$sStringOut = StringReplace($sFileRead, $MyArrayToPatch[$i], $MyArrayToPatch[$i + 1], 0, 1)
$sFileRead = $sStringOut
$sStringOut = $sFileRead
ProgressWrite(Round($i / $iRows * 100))
Next
FileClose($hFileOpen)
FileMove($MyFileToPatch, $MyFileToPatch & ".bak", $FC_OVERWRITE)
Local $hFileOpen1 = FileOpen($MyFileToPatch, $FO_OVERWRITE + $FO_BINARY)
FileWrite($hFileOpen1, Binary($sStringOut))
FileClose($hFileOpen1)
ProgressWrite(0)
Sleep(100)
LogWrite(1, "File patched." & @CRLF)
Else
MemoWrite(@CRLF & "No patterns were found" & @CRLF & "---" & @CRLF & "or" & @CRLF & "---" & @CRLF & "file is already patched.")
Sleep(100)
LogWrite(1, "No patterns were found or file already patched." & @CRLF)
EndIf
EndFunc
Func RestoreFile($MyFileToDelete)
If FileExists($MyFileToDelete & ".bak") Then
FileDelete($MyFileToDelete)
FileMove($MyFileToDelete & ".bak", $MyFileToDelete, $FC_OVERWRITE)
Sleep(100)
MemoWrite(@CRLF & "File restored" & @CRLF & "---" & @CRLF & $MyFileToDelete)
LogWrite(1, $MyFileToDelete)
LogWrite(1, "File restored.")
Else
Sleep(100)
MemoWrite(@CRLF & "No backup file found" & @CRLF & "---" & @CRLF & $MyFileToDelete)
LogWrite(1, $MyFileToDelete)
LogWrite(1, "No backup file found.")
EndIf
EndFunc
Func BlockPopUp()
GUICtrlSetState($idBtnBlockPopUp, 128)
MemoWrite(@CRLF & "Checking for Internet connectivity" & @CRLF & "---" & @CRLF & "Please wait...")
Local $sCmdInfo = "PowerShell Set-ExecutionPolicy Bypass -scope Process -Force;(Get-NetRoute | Where-Object DestinationPrefix -eq '0.0.0.0/0' | Get-NetIPInterface | Where-Object ConnectionState -eq 'Connected') -ne $null"
Local $iPID = Run($sCmdInfo, "", @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
Local $sOutput = ""
While 1
$sOutput &= StdoutRead($iPID)
If @error Then ExitLoop
WEnd
ProcessWaitClose($iPID)
If StringReplace($sOutput, @CRLF, "") = "True" Then
MemoWrite(@CRLF & "Searching for IP Addresses" & @CRLF & "---" & @CRLF & "Please wait...")
$sCmdInfo = "PowerShell Set-ExecutionPolicy Bypass -scope Process -Force;$ips=@();$soa=(Resolve-DnsName -Name adobe.io -Type SOA).PrimaryServer;Do{$ip=(Resolve-DnsName -Name adobe.io -Server $soa).IPAddress;$ips+=$ip;$ips=$ips|Select -Unique|Sort-Object}While($ips.Count -lt 8);$list=$ips -join ',';$list"
$iPID = Run($sCmdInfo, "", @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
$sOutput = ""
While 1
$sOutput &= StdoutRead($iPID)
If @error Then ExitLoop
WEnd
ProcessWaitClose($iPID)
MemoWrite(@CRLF & "Creating Windows Firewall Rule" & @CRLF & "---" & @CRLF & "Blocking:" & @CRLF & StringReplace($sOutput, @CRLF, ""))
$sCmdInfo = "netsh advfirewall firewall delete rule name=""Adobe Unlicensed Pop-up"""
$iPID = Run($sCmdInfo, "", @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
ProcessWaitClose($iPID)
$sCmdInfo = "netsh advfirewall firewall add rule name=""Adobe Unlicensed Pop-up"" dir=out action=block remoteip=""" & StringReplace($sOutput, @CRLF, "") & """"
$iPID = Run($sCmdInfo, "", @SW_HIDE, BitOR($STDERR_CHILD, $STDOUT_CHILD))
ProcessWaitClose($iPID)
MemoWrite(@CRLF & "Added Windows Firewall Rule" & @CRLF & "---" & @CRLF & "Blocking:" & @CRLF & StringReplace($sOutput, @CRLF, ""))
Else
MemoWrite(@CRLF & "No Internet Connectivity" & @CRLF & "---" & @CRLF & "Please check your Internet connection and try again.")
GUICtrlSetState($idBtnBlockPopUp, 64)
EndIf
EndFunc
Func _ListView_LeftClick($hListView, $lParam)
Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
Local $iIndex = DllStructGetData($tInfo, "Index")
If $iIndex <> -1 Then
Local $iX = DllStructGetData($tInfo, "X")
Local $aIconRect = _GUICtrlListView_GetItemRect($hListView, $iIndex, 1)
If $iX < $aIconRect[0] And $iX >= 5 Then
Return 0
Else
Local $aHit
$aHit = _GUICtrlListView_HitTest($g_idListview)
If $aHit[0] <> -1 Then
Local $GroupIdOfHitItem = _GUICtrlListView_GetItemGroupID($idListview, $aHit[0])
If _GUICtrlListView_GetItemChecked($g_idListview, $aHit[0]) = 1 Then
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
If _GUICtrlListView_GetItemGroupID($idListview, $i) = $GroupIdOfHitItem Then
_GUICtrlListView_SetItemChecked($g_idListview, $i, 0)
EndIf
Next
Else
For $i = 0 To _GUICtrlListView_GetItemCount($idListview) - 1
If _GUICtrlListView_GetItemGroupID($idListview, $i) = $GroupIdOfHitItem Then
_GUICtrlListView_SetItemChecked($g_idListview, $i, 1)
EndIf
Next
EndIf
EndIf
EndIf
EndIf
EndFunc
Func _ListView_RightClick()
Local $aHit
$aHit = _GUICtrlListView_HitTest($g_idListview)
If $aHit[0] <> -1 Then
If _GUICtrlListView_GetItemChecked($g_idListview, $aHit[0]) = 1 Then
_GUICtrlListView_SetItemChecked($g_idListview, $aHit[0], 0)
Else
_GUICtrlListView_SetItemChecked($g_idListview, $aHit[0], 1)
EndIf
EndIf
EndFunc
Func _Assign_Groups_To_Found_Files()
Local $MyListItemCount = _GUICtrlListView_GetItemCount($idListview)
Local $ItemFromList
For $i = 0 To $MyListItemCount - 1
_GUICtrlListView_SetItemChecked($idListview, $i)
$ItemFromList = _GUICtrlListView_GetItemText($idListview, $i, 1)
Select
Case StringInStr($ItemFromList, "Acrobat")
_GUICtrlListView_InsertGroup($idListview, $i, 1, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 1)
_GUICtrlListView_SetGroupInfo($idListview, 1, "Acrobat", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Aero")
_GUICtrlListView_InsertGroup($idListview, $i, 2, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 2)
_GUICtrlListView_SetGroupInfo($idListview, 2, "Aero", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "After Effects")
_GUICtrlListView_InsertGroup($idListview, $i, 3, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 3)
_GUICtrlListView_SetGroupInfo($idListview, 3, "After Effects", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Animate")
_GUICtrlListView_InsertGroup($idListview, $i, 4, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 4)
_GUICtrlListView_SetGroupInfo($idListview, 4, "Animate", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Audition")
_GUICtrlListView_InsertGroup($idListview, $i, 5, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 5)
_GUICtrlListView_SetGroupInfo($idListview, 5, "Audition", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Adobe Bridge")
_GUICtrlListView_InsertGroup($idListview, $i, 6, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 6)
_GUICtrlListView_SetGroupInfo($idListview, 6, "Bridge", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Character Animator")
_GUICtrlListView_InsertGroup($idListview, $i, 7, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 7)
_GUICtrlListView_SetGroupInfo($idListview, 7, "Character Animator", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "AppsPanel")
_GUICtrlListView_InsertGroup($idListview, $i, 8, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 8)
_GUICtrlListView_SetGroupInfo($idListview, 8, "Creative Cloud", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Dimension")
_GUICtrlListView_InsertGroup($idListview, $i, 9, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 9)
_GUICtrlListView_SetGroupInfo($idListview, 9, "Dimension", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Dreamweaver")
_GUICtrlListView_InsertGroup($idListview, $i, 10, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 10)
_GUICtrlListView_SetGroupInfo($idListview, 10, "Dreamweaver", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Illustrator")
_GUICtrlListView_InsertGroup($idListview, $i, 11, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 11)
_GUICtrlListView_SetGroupInfo($idListview, 11, "Illustrator", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "InCopy")
_GUICtrlListView_InsertGroup($idListview, $i, 12, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 12)
_GUICtrlListView_SetGroupInfo($idListview, 12, "InCopy", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "InDesign")
_GUICtrlListView_InsertGroup($idListview, $i, 13, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 13)
_GUICtrlListView_SetGroupInfo($idListview, 13, "InDesign", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Lightroom CC")
_GUICtrlListView_InsertGroup($idListview, $i, 14, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 14)
_GUICtrlListView_SetGroupInfo($idListview, 14, "Lightroom CC", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Lightroom Classic")
_GUICtrlListView_InsertGroup($idListview, $i, 15, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 15)
_GUICtrlListView_SetGroupInfo($idListview, 15, "Lightroom Classic", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Media Encoder")
_GUICtrlListView_InsertGroup($idListview, $i, 16, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 16)
_GUICtrlListView_SetGroupInfo($idListview, 16, "Media Encoder", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Photoshop")
_GUICtrlListView_InsertGroup($idListview, $i, 17, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 17)
_GUICtrlListView_SetGroupInfo($idListview, 17, "Photoshop", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Premiere Pro")
_GUICtrlListView_InsertGroup($idListview, $i, 18, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 18)
_GUICtrlListView_SetGroupInfo($idListview, 18, "Premiere Pro", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Premiere Rush")
_GUICtrlListView_InsertGroup($idListview, $i, 19, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 19)
_GUICtrlListView_SetGroupInfo($idListview, 19, "Premiere Rush", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Substance 3D Designer")
_GUICtrlListView_InsertGroup($idListview, $i, 20, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 20)
_GUICtrlListView_SetGroupInfo($idListview, 20, "Substance 3D Designer", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Substance 3D Modeler")
_GUICtrlListView_InsertGroup($idListview, $i, 21, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 21)
_GUICtrlListView_SetGroupInfo($idListview, 21, "Substance 3D Modeler", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Substance 3D Painter")
_GUICtrlListView_InsertGroup($idListview, $i, 22, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 22)
_GUICtrlListView_SetGroupInfo($idListview, 22, "Substance 3D Painter", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Substance 3D Sampler")
_GUICtrlListView_InsertGroup($idListview, $i, 23, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 23)
_GUICtrlListView_SetGroupInfo($idListview, 23, "Substance 3D Sampler", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Substance 3D Stager")
_GUICtrlListView_InsertGroup($idListview, $i, 24, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 24)
_GUICtrlListView_SetGroupInfo($idListview, 24, "Substance 3D Stager", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Adobe.Fresco")
_GUICtrlListView_InsertGroup($idListview, $i, 25, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 25)
_GUICtrlListView_SetGroupInfo($idListview, 25, "Fresco", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "Adobe.XD")
_GUICtrlListView_InsertGroup($idListview, $i, 26, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 26)
_GUICtrlListView_SetGroupInfo($idListview, 26, "XD", 1, $LVGS_COLLAPSIBLE)
Case StringInStr($ItemFromList, "PhotoshopExpress")
_GUICtrlListView_InsertGroup($idListview, $i, 27, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 27)
_GUICtrlListView_SetGroupInfo($idListview, 27, "PhotoshopExpress", 1, $LVGS_COLLAPSIBLE)
Case Else
_GUICtrlListView_InsertGroup($idListview, $i, 28, "", 1)
_GUICtrlListView_SetItemGroupID($idListview, $i, 28)
_GUICtrlListView_SetGroupInfo($idListview, 28, "Else", 1, $LVGS_COLLAPSIBLE)
EndSelect
Next
EndFunc
Func _Collapse_All_Click()
Local $aInfo, $aCount = _GUICtrlListView_GetGroupCount($idListview)
If $aCount > 0 Then
If $MyLVGroupIsExpanded = 1 Then
For $i = 1 To 28
$aInfo = _GUICtrlListView_GetGroupInfo($idListview, $i)
_GUICtrlListView_SetGroupInfo($idListview, $i, $aInfo[0], $aInfo[1], $LVGS_COLLAPSED)
Next
Else
_Expand_All_Click()
EndIf
$MyLVGroupIsExpanded = Not $MyLVGroupIsExpanded
EndIf
EndFunc
Func _Expand_All_Click()
Local $aInfo, $aCount = _GUICtrlListView_GetGroupCount($idListview)
If $aCount > 0 Then
For $i = 1 To 28
$aInfo = _GUICtrlListView_GetGroupInfo($idListview, $i)
_GUICtrlListView_SetGroupInfo($idListview, $i, $aInfo[0], $aInfo[1], $LVGS_NORMAL)
_GUICtrlListView_SetGroupInfo($idListview, $i, $aInfo[0], $aInfo[1], $LVGS_COLLAPSIBLE)
Next
EndIf
EndFunc
Func WM_COMMAND($hWnd, $Msg, $wParam, $lParam)
If BitAND($wParam, 0x0000FFFF) = $idButtonStop Then $fInterrupt = 1
Return $GUI_RUNDEFMSG
EndFunc
Func WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
#forceref $hWnd, $iMsg, $wParam, $lParam
Local $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
Local $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
Local $iCode = DllStructGetData($tNMHDR, "Code")
Switch $hWndFrom
Case $g_idListview
Switch $iCode
Case $LVN_COLUMNCLICK
_Collapse_All_Click()
Case $NM_CLICK
_ListView_LeftClick($g_idListview, $lParam)
Case $NM_RCLICK
_ListView_RightClick()
EndSwitch
EndSwitch
Return $GUI_RUNDEFMSG
EndFunc
Func _Exit()
FileDelete(@WindowsDir & "\Temp\RunAsTI.exe")
Exit
EndFunc
Func IniReadArray($FileName, $section, $key, $default)
Local $sINI = IniRead($FileName, $section, $key, $default)
$sINI = StringReplace($sINI, '"', '')
StringReplace($sINI, ",", ",")
Local $aSize = @extended
Local $aReturn[$aSize + 1]
Local $aSplit = StringSplit($sINI, ",")
For $i = 0 To $aSize
$aReturn[$i] = $aSplit[$i + 1]
Next
Return $aReturn
EndFunc
