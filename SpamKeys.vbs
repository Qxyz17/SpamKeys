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
result = MsgBox("����ժҪ��" & vbCrLf & "��Ϣģ��: " & baseMessage & vbCrLf & "��������: " & iterations & vbCrLf & "�����ٶ�: " & speed & "��/��" & vbCrLf & "���ͷ�ʽ: " & IIf(sendMethod = "1", "Enter", "Ctrl+Enter") & vbCrLf & vbCrLf & "�Ƿ�ʼִ�У�" & vbCrLf & vbCrLf & "��ʾ���� Ctrl+C ����ʱ��ֹˢ��", vbYesNo + vbInformation, "ȷ������")
If result = vbNo Then
    WScript.Quit
End If
Set WshShell = WScript.CreateObject("WScript.Shell")
MsgBox "�ű�����5���ʼ����ȷ��QQ�����Ѽ����ҹ����������У�" & vbCrLf & "��ȷ����������л���QQ���ڡ�" & vbCrLf & "��ʾ���� Ctrl+C ����ʱ��ֹˢ��", vbInformation, "��ʾ"
WScript.Sleep 5000
For i = 1 to CInt(iterations)
    If CheckForBreak() Then
        MsgBox "ˢ������ֹ���ѷ��� " & (i-1) & " ����Ϣ��", vbInformation, "��ֹ"
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
    MsgBox "��Ϣ������ɣ������� " & iterations & " ����Ϣ��", vbInformation, "���"
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