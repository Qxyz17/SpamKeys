Option Explicit
On Error Resume Next
Dim message, baseMessage, randomPart, i, j, randomStr
Dim WshShell, iterations, delay, sendMethod, result, speed, sleepTime
baseMessage = InputBox("请输入要发送的消息内容（使用{random}表示随机字符串）：", "消息设置")
If baseMessage = "" Then
    WScript.Quit
End If
iterations = InputBox("请输入要发送的消息条数：", "发送数量", "10")
If iterations = "" Or Not IsNumeric(iterations) Then
    MsgBox "无效的输入，脚本将退出。"
    WScript.Quit
End If
speed = InputBox("请输入每秒发送的消息条数（1-1000）：", "发送速度", "10")
If speed = "" Or Not IsNumeric(speed) Then
    MsgBox "无效的输入，脚本将退出。"
    WScript.Quit
End If
speed = CInt(speed)
If speed < 1 Then speed = 1
If speed > 1000 Then speed = 1000
sleepTime = 1000 / speed
sendMethod = InputBox("请选择发送方式：" & vbCrLf & "1 - 按Enter发送" & vbCrLf & "2 - 按Ctrl+Enter发送", "发送方式", "1")
If sendMethod = "" Then
    WScript.Quit
End If
If sendMethod <> "1" And sendMethod <> "2" Then
    MsgBox "无效的选择，脚本将退出。"
    WScript.Quit
End If
result = MsgBox("设置摘要：" & vbCrLf & "消息模板: " & baseMessage & vbCrLf & "发送条数: " & iterations & vbCrLf & "发送速度: " & speed & "条/秒" & vbCrLf & "发送方式: " & IIf(sendMethod = "1", "Enter", "Ctrl+Enter") & vbCrLf & vbCrLf & "是否开始执行？" & vbCrLf & vbCrLf & "提示：按 Ctrl+C 可随时中止刷屏", vbYesNo + vbInformation, "确认设置")
If result = vbNo Then
    WScript.Quit
End If
Set WshShell = WScript.CreateObject("WScript.Shell")
MsgBox "脚本将在5秒后开始，请确保QQ窗口已激活且光标在输入框中！" & vbCrLf & "按确定后请快速切换到QQ窗口。" & vbCrLf & "提示：按 Ctrl+C 可随时中止刷屏", vbInformation, "提示"
WScript.Sleep 5000
For i = 1 to CInt(iterations)
    If CheckForBreak() Then
        MsgBox "刷屏已中止！已发送 " & (i-1) & " 条消息。", vbInformation, "中止"
        Exit For
    End If
    randomStr = ""
    For j = 1 to 5
        If Rnd() > 0.5 Then
            randomStr = randomStr & Chr(97 + CInt(Rnd() * 25))
        Else
            randomStr = randomStr & CInt(Rnd() * 9)
        End If
    Next
    message = Replace(baseMessage, "{random}", randomStr)
    WshShell.SendKeys EncodeSendKeys(message)
    If sendMethod = "1" Then
        WshShell.SendKeys "{ENTER}"
    Else
        WshShell.SendKeys "^({{ENTER})"
    End If
    WScript.Sleep sleepTime
Next
If i > CInt(iterations) Then
    MsgBox "消息发送完成！共发送 " & iterations & " 条消息。", vbInformation, "完成"
End If
Set WshShell = Nothing
Function IIf(expr, trueVal, falseVal)
    If expr Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function
Function EncodeSendKeys(text)
    Dim i, result
    result = ""
    For i = 1 to Len(text)
        Dim currentChar
        currentChar = Mid(text, i, 1)
        Select Case currentChar
            Case "+", "^", "%", "~", "(", ")", "{", "}", "[", "]", "\"
                result = result & "{" & currentChar & "}"
            Case Else
                result = result & currentChar
        End Select
    Next
    EncodeSendKeys = result
End Function
Function CheckForBreak()
    Dim breakKey
    breakKey = &H43
    CheckForBreak = (GetAsyncKeyState(breakKey) And &H8000) <> 0
End Function
Function GetAsyncKeyState(vKey)
    Dim WshShell, result
    Set WshShell = CreateObject("WScript.Shell")
    result = WshShell.Run("powershell -Command ""Add-Type -TypeDefinition '[DllImport(""user32.dll"")]public static extern short GetAsyncKeyState(int vKey);' -Name WinAPI -Namespace Internal; [Internal.WinAPI]::GetAsyncKeyState(" & vKey & ")", 0, True)
    GetAsyncKeyState = result
End Function