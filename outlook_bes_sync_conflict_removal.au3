#NoTrayIcon

#include <Array.au3>
#include <Timers.au3>

Global $oMyError = ObjEvent("AutoIt.Error", "ErrFunc")
Global $lastRemoval = TimerInit()
Global $removeReady = False
Global $receiptSetting = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Options\Mail","Receipt Response")

Global $firstStart = True
Global $idleTime

const $olFolderConflicts=19
const $olFolderSyncIssues=20
const $olFolderDeletedItems=3
const $PR_SEARCH_KEY = "http://schemas.microsoft.com/mapi/proptag/0x300B0102"

AdlibRegister("_idleControl", 500)

While 1
    Sleep(100)
    If $removeReady = True Then
        If TimerDiff($lastRemoval) > (10 * 60 * 1000) Or $firstStart = True Then ;more than 10 minutes
            $firstStart = False
            _clearSyncsConfs()
        EndIf
    EndIf
WEnd

Func _clearSyncsConfs()
    Dim $searchKeys[1]
    $searchKeys[0] = 0
    $oOApp = ObjCreate("Outlook.Application")
    If $oOApp = 0 Then Exit;looks like Outlook isn't installed
    $myNamespace = $oOApp.GetNamespace("MAPI")

    ;store the current receipt response setting
    $receiptSetting = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Options\Mail","Receipt Response")
    ;don't respond to any possible receipts, as these are duplicates we're deleting
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Options\Mail", "Receipt Response", "REG_DWORD", 1)

    ;empty conflicts
    $myFolder = $myNamespace.GetDefaultFolder($olFolderConflicts)
    $myItems = $myFolder.Items().Count
    For $i = $myItems To 1 Step -1
        ;get the items search key which remains unique and does not change
        _ArrayAdd($searchKeys, $myFolder.Items($i).PropertyAccessor.GetProperty($PR_SEARCH_KEY))
        ;mark as read first to prevent read receipt confusion massacre
        $myFolder.Items($i).UnRead = False
        ;delete
        $myFolder.Items($i).Delete()

        If $removeReady = False Then
            _stopDeletion()
            _idleWait()
        EndIf
    Next

    ;empty sync issues
    $myFolder = $myNamespace.GetDefaultFolder($olFolderSyncIssues)
    $myItems = $myFolder.Items().Count
    For $i = $myItems To 1 Step -1
        _ArrayAdd($searchKeys, $myFolder.Items($i).PropertyAccessor.GetProperty($PR_SEARCH_KEY))
        $myFolder.Items($i).UnRead = False
        $myFolder.Items($i).Delete()

        If $removeReady = False Then
            _stopDeletion()
            _idleWait()
        EndIf
    Next

    ;permanently delete sync issues and conflicts
    $myFolder = $myNamespace.GetDefaultFolder($olFolderDeletedItems)
    $myItems = $myFolder.Items().Count
    For $i = $myItems To 1 Step -1
        $found = _ArraySearch($searchKeys, $myFolder.Items($i).PropertyAccessor.GetProperty($PR_SEARCH_KEY), 1)
        If $found > 0 Then $myFolder.Items($i).Delete()

        If $removeReady = False Then
            _stopDeletion()
            _idleWait()
        EndIf
    Next

    ;restore the receipt response setting
    _stopDeletion()

    $lastRemoval = TimerInit()

    Return
EndFunc

Func _idleControl()
    $idleTime = _Timer_GetIdleTime()
    If $idleTime > (60 * 1000) Then ; 1 minute
        $removeReady = True
    Else
        $removeReady = False
    EndIf
EndFunc

Func _idleWait()
    While $removeReady = False
        Sleep(1000)
    WEnd
    $receiptSetting = RegRead("HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Options\Mail","Receipt Response")
EndFunc

Func _stopDeletion()
    Sleep(3500)
    RegWrite("HKEY_CURRENT_USER\Software\Microsoft\Office\12.0\Outlook\Options\Mail", "Receipt Response", "REG_DWORD", $receiptSetting)
EndFunc

Func ErrFunc()
    AdlibUnRegister("_idleControl")
    _stopDeletion()
    $lastRemoval = TimerInit()
    $removeReady = False

    $error = "We intercepted an Error !"      & @CRLF  & @CRLF & _
             "err.description is: "    & @TAB & $oMyError.description    & @CRLF & _
             "err.windescription:"     & @TAB & $oMyError.windescription & @CRLF & _
             "err.number is: "         & @TAB & hex($oMyError.number,8)  & @CRLF & _
             "err.lastdllerror is: "   & @TAB & $oMyError.lastdllerror   & @CRLF & _
             "err.scriptline is: "     & @TAB & $oMyError.scriptline     & @CRLF & _
             "err.source is: "         & @TAB & $oMyError.source         & @CRLF & _
             "err.helpfile is: "       & @TAB & $oMyError.helpfile       & @CRLF & _
             "err.helpcontext is: "    & @TAB & $oMyError.helpcontext

    MsgBox(0, "Error Removing Items from Outlook", $error)

    Local $err = $oMyError.number
    If $err = 0 Then $err = -1

    $g_eventerror = $err  ; to check for after this function returns

    Exit
Endfunc