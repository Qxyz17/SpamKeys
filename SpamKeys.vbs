Option Explicit
On Error Resume Next

Dim message, baseMessage, randomPart, i, j, randomStr
Dim WshShell, iterations, delay, sendMethod, result, speed, sleepTime

baseMessage = InputBox("������Ҫ���͵���Ϣ���ݣ�ʹ��{random}��ʾ����ַ�������", "��Ϣ����")

If baseMessage = "" Then
    WScript.Quit
End If

iterations = InputBox("������Ҫ���͵���Ϣ������", "��������", "10")

If iterations = "" Or Not IsNumeric(iterations) Then
    MsgBox "��Ч�����룬�ű����˳���"
    WScript.Quit
End If

speed = InputBox("������ÿ�뷢�͵���Ϣ������1-1000����", "�����ٶ�", "10")

If speed = "" Or Not IsNumeric(speed) Then
    MsgBox "��Ч�����룬�ű����˳���"
    WScript.Quit
End If

speed = CInt(speed)
If speed < 1 Then speed = 1
If speed > 1000 Then speed = 1000

sleepTime = 1000 / speed

sendMethod = InputBox("��ѡ���ͷ�ʽ��" & vbCrLf & "1 - ��Enter����" & vbCrLf & "2 - ��Ctrl+Enter����", "���ͷ�ʽ", "1")

If sendMethod = "" Then
    WScript.Quit
End If

If sendMethod <> "1" And sendMethod <> "2" Then
    MsgBox "��Ч��ѡ�񣬽ű����˳���"
    WScript.Quit
End If

result = MsgBox("����ժҪ��" & vbCrLf & _
                "��Ϣģ��: " & baseMessage & vbCrLf & _
                "��������: " & iterations & vbCrLf & _
                "�����ٶ�: " & speed & "��/��" & vbCrLf & _
                "���ͷ�ʽ: " & IIf(sendMethod = "1", "Enter", "Ctrl+Enter") & vbCrLf & vbCrLf & _
                "�Ƿ�ʼִ�У�", vbYesNo + vbInformation, "ȷ������")

If result = vbNo Then
    WScript.Quit
End If

Set WshShell = WScript.CreateObject("WScript.Shell")

MsgBox "�ű�����5���ʼ����ȷ��QQ�����Ѽ����ҹ����������У�" & vbCrLf & "��ȷ����������л���QQ���ڡ�", vbInformation, "��ʾ"
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

MsgBox "��Ϣ������ɣ������� " & iterations & " ����Ϣ��", vbInformation, "���"
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