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

result = MsgBox("设置摘要：" & vbCrLf & _
                "消息模板: " & baseMessage & vbCrLf & _
                "发送条数: " & iterations & vbCrLf & _
                "发送速度: " & speed & "条/秒" & vbCrLf & _
                "发送方式: " & IIf(sendMethod = "1", "Enter", "Ctrl+Enter") & vbCrLf & vbCrLf & _
                "是否开始执行？", vbYesNo + vbInformation, "确认设置")

If result = vbNo Then
    WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

MsgBox "脚本将在5秒后开始，请确保QQ窗口已激活且光标在输入框中！" & vbCrLf & "按确定后请快速切换到QQ窗口。", vbInformation, "提示"
WScript.Sleep 5000

For i = 1 to CInt(iterations)
    randomStr = ""
    For j = 1 to 5
        If Rnd() > 0.5 Then
            randomStr = randomStr & Chr(97 + CInt(Rnd() * 25))
        Else
            randomStr = randomStr & CInt(Rnd() * 9)
        End If
    Next
    
    message = Replace(baseMessage, "{random}", randomStr)
    
    SendTextWithClipboard WshShell, message
    
    If sendMethod = "1" Then
        WshShell.SendKeys "{ENTER}"
    Else
        WshShell.SendKeys "^({{ENTER})"
    End If
    
    WScript.Sleep sleepTime
Next

MsgBox "消息发送完成！共发送 " & iterations & " 条消息。", vbInformation, "完成"
Set WshShell = Nothing

Function IIf(expr, trueVal, falseVal)
    If expr Then
        IIf = trueVal
    Else
        IIf = falseVal
    End If
End Function

Sub SendTextWithClipboard(shellObj, textToSend)
    Dim clipBoard
    Set clipBoard = CreateObject("HTMLFile")
    
    On Error Resume Next
    Dim oldClipboard
    oldClipboard = shellObj.Exec("mshta.exe ""javascript:clipboardData.getData('Text');close();""").StdOut.ReadAll
    
    shellObj.Run "mshta.exe ""javascript:clipboardData.setData('Text','" & Replace(textToSend, "'", "\'") & "');close();""", 0, True
    WScript.Sleep 100
    
    shellObj.SendKeys "^v"
    WScript.Sleep 100
    
    If Err.Number = 0 And oldClipboard <> "" Then
        shellObj.Run "mshta.exe ""javascript:clipboardData.setData('Text','" & Replace(oldClipboard, "'", "\'") & "');close();""", 0, True
    End If
    On Error GoTo 0
End Sub