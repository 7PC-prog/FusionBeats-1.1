' Просто запускает HTA
Option Explicit
Dim shell
Set shell = CreateObject("WScript.Shell")
shell.Run "FloppotronGUI.hta", 1, False
