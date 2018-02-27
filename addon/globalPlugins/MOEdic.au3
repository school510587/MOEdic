#NoTrayIcon
#pragma compile(UPX, True)
#pragma compile(LegalCopyright, "(C) 2017-2018 Coscell Kao and Bo-Cheng Jhan")
#pragma compile(FileVersion, 2.7.7)
#pragma compile(FileDescription, 國字快查)
#include <Array.au3>
#include <Clipboard.au3>
#include <IE.au3>
#include <StringConstants.au3>
#include <WinAPI.au3>
#include <WinAPICom.au3>
#include <WinAPIShPath.au3>

Global Const $CHINESE_RANGE = "\x{2E80}-\x{2FDF}\x{3400}-\x{9FFF}\x{F900}-\x{FAFF}"

Global $query = "", $quick = ($CmdLine[0] > 0 And $CmdLine[1] == "/quick")
If Not HotKeySet("^#q", "NullHotKeyHandler") Then
    MsgBox($MB_OK, "錯誤", "無法設置快速鍵 Ctrl + Win + Q, 可能有別的程序正在使用它。")
    Exit 1
EndIf
If _ClipBoard_IsFormatAvailable($CF_TEXT) Then
    ConsoleWriteLine("嘗試從剪貼簿取得輸入。", True)
    $query = ClipGet()
    If StringLen($query) > 20 Then $query = ""
EndIf
While 1
    If Not $quick Then
        $query = InputBox("教育部重編國語辭典修訂本", "我想找", $query)
        If @error Then Exit 1; Execution is cancelled.
    EndIf
    If StringRegExp($query, "^([" & $CHINESE_RANGE & "\x{FF0C}]{1,20}|[\-A-Za-z]{1,20})$") Then ExitLoop
    If Not $quick Then MsgBox($MB_OK, "錯誤", "查詢詞必須是不超過 20 字的純中文（含逗號）或純英文字串。")
    $quick = False; Show prompt if the content of clipboard is not valid.
WEnd
ConsoleWriteLine("輸入結果：" & $query & "。", True)

Message("開始查詢：" & $query)
_IELoadWaitTimeout(10000); 10 seconds.
Global $oIEFace = Null, $oIEWorker = Null
OnAutoItExitRegister("Clean")
Global $data = "", $total = 0, $progress = 0, $loading = True, $pause = False
HotKeySet("^#q", "SwitchPause")
Local $msg = FindQueryResult($oIEWorker, $query, 0)
If @error Then
    Local $e = @error
    HotKeySet("^#q")
    HideHintText()
    MsgBox($MB_OK, $e < 0 ? "結束" : "錯誤", $msg)
    Exit -Int(($e <> -1))
EndIf
$loading = False
$total = @extended

Local $h = 0
While 1
    NavigateMoreResults($oIEFace, $oIEWorker)
    If Not IsObj($oIEWorker) And $oIEWorker <> Default Then ExitLoop
    $face_hwnd = Execute("_IEPropertyGet($oIEFace, 'hwnd')")
    If Not @error And WinWaitClose($face_hwnd, "", 1) Then
        HotKeySet("^#q", "ShowHintText")
        ExitLoop
    EndIf
WEnd
$h = Execute("_IEPropertyGet($oIEFace, 'hwnd')")
If Not @error Then WinWaitClose($h); Use $h to avoid waiting for closing an empty-title window.

Func Clean()
    HideHintText(); Hide messages anyway.
    If IsObj($oIEWorker) Then _IEQuit($oIEWorker)
    If IsObj($oIEFace) Then _IEQuit($oIEFace)
    If @exitCode < 0 Then ConsoleWriteLine("執行中有發生錯誤，請檢查事件記錄。")
EndFunc

Func ConsoleWriteLine($line, $stdout=False, $encoding=$SB_UTF8)
    $line = StringToBinary($line & @CRLF, $encoding)
    If $stdout Then Return ConsoleWrite($line)
    Return ConsoleWriteError($line)
EndFunc

Func FindQueryResult(ByRef $oIE, $query, $index)
    $oIE = _IECreate("http://dict.revised.moe.edu.tw/cbdic/search.htm", 0, 0, 0)
    If Not IsObj($oIE) Then Return SetError(1, 0, "無法開啟教育部字典網頁。")
    $oIE.Silent = 1
    LoadWait($oIE)
    If @error Then Return SetError(-1, 0, "操作已被用戶中斷。")
    Local $oBody = _IEDocGetObj($oIE).body, $oQuery
    For $i = 0 To 1; It is limited to try twice.
        $oQuery = _IEGetObjById($oBody, "qs0")
        If Not @error And IsObj($oQuery) Then ExitLoop; The place to put query text is found.
        If $i > 0 Then Return SetError(2, 0, "無法定位查詢輸入點。")
        $oQuery = _IETagNameGetCollection($oBody, "input")
        If Not @error And IsObj($oQuery) Then
            For $b In $oQuery
                If $b.value = "本典" Then
                    _IEAction($b, "click")
                    LoadWait($oIE)
                    If @error Then Return SetError(-1, 0, "操作已被用戶中斷。")
                    $oBody = _IEDocGetObj($oIE).body
                EndIf
            Next
        EndIf
    Next
    _IEFormElementSetValue($oQuery, $query)
    $oQuery = $oQuery.form; The <form> element containing it.
    _IEFormSubmit($oQuery, 0)
    LoadWait($oIE)
    If @error Then Return SetError(-1, 0, "操作已被用戶中斷。")
    $oBody = _IEDocGetObj($oIE).body
    Local $tmp = _IEBodyReadHTML($oIE)
    If StringInStr($tmp, "檢索結果列表：") > 0 Then; Multiple results.
        ConsoleWriteLine("找到多筆檢索結果。", True)
        Local $oBold = _IETagNameGetCollection($oBody, "b"), $total
        For $o In $oBold
            $o = StringRegExp($o.innerText, "(*UCP)共\s*(\d+)\s*筆", $STR_REGEXPARRAYMATCH)
            If Not @error And IsArray($o) Then
                $total = Int($o[0])
                ExitLoop
            EndIf
        Next
        If $total <= 0 Then Return SetError(3, 0, "沒有找到資料總數。")
        ConsoleWriteLine("檢索結果總共 " & $total & " 筆。", True)
        If $index >= $total Then
            ConsoleWriteLine("要求的檢索結果 [" & $index & "] 不存在。", True)
            Return SetError(4, 0, "所指定的結果資料不存在。")
        EndIf
        Local $p = _IEGetObjById($oBody, "psize")
        If @error Or $p.tagName <> "select" Then
            If @error Then
                ConsoleWriteLine('找不到每頁資料數目設定 (id="psize").', True)
            Else
                ConsoleWriteLine("每頁資料數目設定的格式不如預期。", True)
                ConsoleWriteLine(_IEPropertyGet($p, "outerhtml"), True)
            EndIf
            Return SetError(5, 0, "找不到每頁資料數目設定。")
        EndIf
        $p = Int($p.options.item($p.selectedIndex).value); Get the selected value.
        If $index >= $p Then; The target page is not page 1.
            Local $l = _IETagNameGetCollection($oIE, "a"), $pg = 1 + Floor($index / $p)
            For $a In $l; Find the link containing page number as a parameter.
                If _IEPropertyGet($a, "innertext") = "[1]" Then
                    $l = $a.href; The link to the first page.
                    ExitLoop
                EndIf
            Next
            ConsoleWriteLine("要跳至第 " & $pg & " 頁。", True)
            If Not IsString($l) Then Return SetError(6, 0, "未找到跳頁的廉潔。")
            $l = StringRegExpReplace($l, "(?<=[?&]jmpage=)\d+", $pg)
            If @error Then
                ConsoleWriteLine("跳頁網址格式不如預期 <" & $l & ">.", True)
                Return SetError(7, 0, "跳頁網址格式不如預期。")
            EndIf
            _IENavigate($oIE, $l, 0); Jump to the specified page.
            LoadWait($oIE)
            If @error Then Return SetError(-1, 0, "操作已被用戶中斷。")
            $oBody = _IEDocGetObj($oIE).body
            $index = Mod($index, $p)
        EndIf; The worker is directed to the correct page now.
        Local $oTable = _IETableGetCollection($oBody, 0)
        If $oTable.rows.length() < $index + 2 Then
            ConsoleWriteLine("檢索結果列表的列數目不合預期 (.rows.length() 傳回 " & $oTable.rows.length() & ").", True)
            Return SetError(8, $oTable.rows.length(), "檢索結果網頁格式不如預期。")
        ElseIf $oTable.rows.item(1).cells.length() < 2 Then
            ConsoleWriteLine("檢索結果列表的欄數目不合預期 (.cells.length() 傳回 " & $oTable.rows.item(1).cells.length() & ").", True)
            Return SetError(9, $oTable.rows.item(1).cells.length(), "檢索結果網頁格式不如預期。")
        EndIf
        Local $o = $oTable.rows.item(1 + $index).cells.item(1)
        $o = _IETagNameGetCollection($o, "a")
        If $o.length() <> 1 Then
            ConsoleWriteLine("無法確實定位檢索結果的第一筆資料，因為找到 " & $o.length() & " 個連結。", True)
            Return SetError(10, $o.length(), "檢索結果網頁格式不如預期。")
        EndIf
        ConsoleWriteLine("抓取第一筆結果 " & $o.item(0).href, True)
        _IENavigate($oIE, $o.item(0).href, 0)
        LoadWait($oIE)
        If @error Then Return SetError(-1, 0, "操作已被用戶中斷。")
        Return SetError(0, $total, "")
    ElseIf StringInStr($tmp, "字級設定") > 0 Then; One single result.
        ConsoleWriteLine("找到單一檢索結果。", True)
        If $index >= 1 Then
            ConsoleWriteLine("要求的檢索結果 [" & $index & "] 不存在。", True)
            Return SetError(-2, $index, "參數錯誤：檢索結果的索引數目。")
        EndIf
        Return SetError(0, 1, "")
    EndIf
    Return SetError(-3, 0, "查無資料。"); No result.
EndFunc

Func GetSingleResult(ByRef $oIE)
    Local $oTable = _IETableGetCollection($oIE, 4)
    If @error Or Not IsObj($oTable) Then Return SetError(1, 0, "找不到字詞資料所在的表格。")
    Local $oInputs = _IETagNameGetCollection($oIE, "input")
    If @error Or Not IsObj($oInputs) Then Return SetError(2, 0, "找不到「本頁網址」所在的編輯區，因為網頁上無 <input> 元素。")
    Local $aTableData = _IETableWriteToArray($oTable), $b = True
    For $o In $oInputs
        If StringLeft($o.value, 4) == "http" Then
            $aTableData[1][0] = StringFormat("<a href='%s' target='_blank'>%s</a>", $o.value, $aTableData[1][0])
            $b = False
            ExitLoop
        EndIf
    Next
    If $b Then Return SetError(3, 0, "找不到「本頁網址」所在的編輯區。")
    For $i = 0 To UBound($aTableData, $UBOUND_COLUMNS) - 1
        $aTableData[1][$i] = StringRegExpReplace($aTableData[1][$i], "(*UCP)^\s+|\s+$", "")
        If $aTableData[1][$i] == "" Then ContinueLoop
        If $i > 0 Then $aTableData[1][$i] = StringRegExpReplace($aTableData[0][$i], "(*UCP)^\s+|\s+$", "") & "：" & $aTableData[1][$i]
        $aTableData[1][$i] &= "<br/>" & @CRLF
    Next
    $b = _ArrayToString($aTableData, "", 1, 1)
    Return SetError(Int($b == ""), 0, $b)
EndFunc

Func HideHintText()
    AdlibUnRegister("ShowHintText")
    AdlibUnRegister("HideHintText")
    SplashOff()
EndFunc

Func KillIE(ByRef $oIE, $result)
    _IEQuit($oIE)
    $oIE = $result
EndFunc

Func LoadWait(ByRef $oIE)
    Do; Interruptible wa`ting.
        _IELoadWait($oIE, 0, 500)
        If Not @error Then Return SetError(@error, @extended, 0)
    Until $pause
    Return SetError($_IEStatus_LoadWaitTimeout, 0, 0)
EndFunc

Func Message($msg)
    AdlibUnRegister("ShowHintText")
    AdlibUnRegister("HideHintText")
    ConsoleWriteLine($msg)
    SplashTextOn($msg, "")
    AdlibRegister("HideHintText", 8000)
EndFunc

Func NavigateMoreResults(ByRef $oIE, ByRef $ie)
    Static $state = TimerInit()
    If Not IsObj($ie) And $ie <> Default Then Return SetError(-1, 0, 0)
    If $pause Then
        If IsObj($ie) And $state <> Null Then
            Local $idle_time = TimerDiff($state)
            If $idle_time > 25 * 60000 Then; The worker has been idle for over 25 minutes.
                KillIE($ie, Default); The worker is killed because of too long idle time.
                Return SetError(0, 0, 0)
            EndIf
            _IEAction($ie, "invisible"); Avoid unexpected state change.
            $ie.Silent = 1
        EndIf
        If $data Then Return SetError(0, 0, 0)
    ElseIf $ie = Default Then; Try to restore the previous progress.
        Local $msg, $e
        $loading = True
        While 1
            $msg = FindQueryResult($ie, $query, $progress)
            If Not @error Then ExitLoop
            $e = @error
            If IsObj($ie) Then KillIE($ie, Default)
            If $e = -1 Or MsgBox($MB_RETRYCANCEL, "錯誤", $msg) = $IDCANCEL Then
                If IsObj($oIE) Then _IEAction($oIE, "visible")
                $pause = True; This flag is ensured to be reset.
                ExitLoop
            EndIf
        WEnd
        $loading = False
        If Not IsObj($ie) Then Return SetError(0, 0, 0)
    EndIf
    Local $oBody, $page_number = 1, $page_size = 0, $tmp
    If IsObj($oIE) Then; The worker thread is currently idle.
        $page_number = _IEGetObjById($oIE, "page-number")
        If @error Then
            ConsoleWriteLine('找不到頁碼元素 (id="page-number").', True)
            $page_number = 1
        Else
            $page_number = Int(_IEPropertyGet($page_number, "innerhtml"))
        EndIf
        $page_size = _IEGetObjById($oIE, "page-size")
        If @error Then
            ConsoleWriteLine('找不到每頁大小元素 (id="page-size").', True)
            $page_size = 0
        Else
            $page_size = $page_size.selectedIndex
        EndIf
        KillIE($oIE, Null); The state of the worker thread becomes busy.
        _IEAction($ie, "invisible")
    EndIf
    Do; Fetch data until $pause flag is raised or no next item.
        _IELoadWait($ie, 0, 500)
        If Not @error Then
            $oBody = _IEDocGetObj($ie).body
            $tmp = GetSingleResult($oBody)
            If @error Then
                $pause = True
                ConsoleWriteLine(StringFormat("第 %d 筆資料抓取失敗：%s\r\n抓取資料已暫停。", $progress + 1, $tmp), True)
                KillIE($ie, Default); The worker is killed because of failed data retrieval.
                ExitLoop; The face IE object is created upon death of the worker.
            EndIf
            $progress += 1
            $data &= StringFormat("<li>%d. %s</li>\r\n", $progress, $tmp)
            $tmp = _IEGetObjById($oBody, "gonext")
            If @error Or Not IsObj($tmp) Then; No next item.
                KillIE($ie, Null); The worker is killed because of no next result.
                HotKeySet("^#q", "ShowHintText")
                ConsoleWriteLine("抓取資料結束，因為沒找到「下一筆」按鈕。", True)
                ExitLoop; The face IE object is created upon death of the worker.
            EndIf
            _IEAction($tmp, "click")
        EndIf
    Until $pause
    $state = TimerInit(); The worker object becomes idle now, but $pause switch is still locked.
    Local $oIE_tmp = _IECreate("about:blank", 0, 0)
    If @error Or Not IsObj($oIE_tmp) Then
        ConsoleWriteLine("建立 Internet Explorer 物件失敗，錯誤代碼 @error = " & @error & ".", True)
        MsgBox($MB_OK, "錯誤", "無法開啟 Internet Explorer 介面，程式（或附加元件）即將停止執行。")
        Exit -1
    EndIf
    ConsoleWriteLine("開始建立簡易查詢結果網頁，嘗試開啟 output.htm.", True)
    Local $html = FileRead("output.htm")
    If @error Then; Use the attached version.
        ConsoleWriteLine("嘗試開啟內建的 output.htm.", True)
        $tmp = _WinAPI_PathAppend(@TempDir, _WinAPI_CreateGUID())
        If Not FileInstall("output.htm", $tmp) Then
            ConsoleWriteLine("無法讀取內建簡易查詢結果網頁格式 output.htm.", True)
            MsgBox($MB_OK, "錯誤", "無法讀取簡易查詢結果網頁格式，" & (@Compiled ? "程式（或附加元件）" : "腳本") & "即將停止執行。")
            Exit -1
        EndIf
        $html = FileRead($tmp)
        FileDelete($tmp)
    EndIf
    _IEDocWriteHTML($oIE_tmp, $html)
    If Not @error And $data Then
        $tmp = _IEGetObjById($oIE_tmp, "result-list")
        _IEPropertySet($tmp, "innerhtml", $data)
        $tmp = _IEDocGetObj($oIE_tmp)
        Select; Try to call "initialize" via "execScript" and "eval" methods.
          Case Not (Execute('$tmp.parentWindow.execScript("initialize(" & $page_number & "," & $page_size & ")")') ? @error : @error)
          Case Not (Execute('$tmp.parentWindow.eval("initialize(" & $page_number & "," & $page_size & ")")') ? @error : @error)
          Case Else
            ConsoleWriteLine("無法在 Internet Explorer 執行 Javascript 初始化頁面。", True)
        EndSelect
    EndIf
    $oIE_tmp.AddressBar = 0
    $oIE_tmp.MenuBar = 0
    $oIE_tmp.ToolBar = 0
    If Not IsObj($ie) Then WinSetState(_IEPropertyGet($oIE_tmp, "hwnd"), "", @SW_MINIMIZE)
    _IEAction($oIE_tmp, "visible")
    If IsObj($ie) Then; Show the face window and set the state of the worker again.
        WinActivate(_IEPropertyGet($oIE_tmp, "hwnd"))
        _IEAction($ie, "invisible")
        $ie.Silent = 1
    Else; Play a sound to remind the user, with the face window minimized.
        _WinAPI_MessageBeep(5)
    EndIf
    $oIE = $oIE_tmp; The hot-key handler can detect that the worker thread becomes idle.
    Return SetError(Int(IsObj($ie)) - 1, 0, 0)
EndFunc

Func NullHotKeyHandler(); A dedicated blank function to test HotKeySet success.
EndFunc

Func ShowHintText()
    Local $msg
    If $total = 0 Then; The system must be in loading state.
        $msg = "正在開啟「教育部重編國語辭典修訂本」網頁。"
    ElseIf $loading Then
        $msg = "正在重新開啟「教育部重編國語辭典修訂本」網頁。"
    Else
        $msg = StringFormat("%d/%d ", $progress, $total)
        If $oIEWorker = Null Then
            $msg &= "已經停止抓取資料。"
        ElseIf $pause Then
            $msg &= "已經暫停抓取資料"
            If $oIEWorker = Default Then $msg &= "，因為閒置較久或之前的錯誤，所以繼續抓取資料前必須重新開啟「教育部重編國語辭典修訂本」網頁"
            $msg &= "。"
        Else
            $msg &= "正在抓取資料。"
        EndIf
    EndIf
    Message($msg)
EndFunc

Func SwitchPause()
    Static $count = 1, $last = Null
    If $last <> Null Then $count = (TimerDiff($last) > 500 ? 1 : ($count + 1))
    $last = TimerInit()
    Switch $count
      Case 1; Show the system state after about 0.26 second.
        AdlibRegister("ShowHintText", 260)
      Case 2; Pause or continue data fetching.
        AdlibUnRegister("ShowHintText")
        If $pause Then; The worker thread is assumed to be paused.
            If Not $loading And IsObj($oIEFace) And WinActive(_IEPropertyGet($oIEFace, "hwnd")) Then
                _IEAction($oIEFace, "invisible")
                _IEAction($oIEWorker, "invisible")
                $pause = False
            Else; The state of the worker thread is not synchronous.
                _WinAPI_MessageBeep()
            EndIf
        Else; The worker thread is assumed to be busy.
            If Not $loading And IsObj($oIEFace) Then; The state of the worker thread is not synchronous.
                _WinAPI_MessageBeep()
            Else; The user asks for viewing the data collected currently.
                $pause = True
            EndIf
        EndIf
      Case Else; Overwrite $count value to avoid overflow.
        $count = 2
    EndSwitch
EndFunc
