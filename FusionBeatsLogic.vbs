' FloppotronLogic.vbs
Option Explicit

Dim args
Set args = WScript.Arguments

If args.Count < 2 Then
    WScript.Echo "Использование: cscript FloppotronLogic.vbs <ritm.txt> <devices.txt>"
    WScript.Quit 1
End If

Dim rhythmFile, devicesStr, fso, rhythmText, devicesArr
rhythmFile = args(0)
devicesStr = args(1)

Set fso = CreateObject("Scripting.FileSystemObject")
If Not fso.FileExists(rhythmFile) Then
    WScript.Echo "Файл ритма не найден: " & rhythmFile
    WScript.Quit 1
End If

rhythmText = fso.OpenTextFile(rhythmFile, 1).ReadAll
devicesArr = Split(devicesStr, ";") ' каждый девайс = Name,Freq,Dur,Delay

Dim shell
Set shell = CreateObject("WScript.Shell")

Dim devicesDict, i, parts
Set devicesDict = CreateObject("Scripting.Dictionary")
For i = 0 To UBound(devicesArr)
    parts = Split(devicesArr(i), ",")
    If UBound(parts) = 3 Then
        devicesDict.Add parts(0), Array(CLng(parts(1)), CLng(parts(2)), CLng(parts(3)))
    End If
Next

Dim hddToolPath
hddToolPath = "hdd_beep.exe" ' <-- ваш инструмент

Dim events, j, deviceName, freq, dur, delay, cmd, subParts
events = Split(rhythmText, " ")

For i = 0 To UBound(events)
    subParts = Split(events(i), "-")
    For j = 0 To UBound(subParts)
        deviceName = subParts(j)
        If devicesDict.Exists(deviceName) Then
            freq = devicesDict(deviceName)(0)
            dur = devicesDict(deviceName)(1)
            delay = devicesDict(deviceName)(2)
            
            cmd = """" & hddToolPath & """" & " """ & deviceName & """ " & freq & " " & dur
            WScript.Echo "Вызов: " & cmd
            shell.Run cmd, 0, False
            WScript.Sleep delay
        End If
    Next
    WScript.Sleep 1000
Next

WScript.Echo "=== Ритм проигран ==="
