<div align="center">

## VB Code to Shell and Wait


</div>

### Description

Executes a command passed as a string, and waits for it to finish. Example: lResult = ShellAndWait("d:\ztbold\ztw.exe", 10000)

That calls the d:\ztbold\ztw.exe and waits for it to exit for up to 10,000 milliseconds.

If you want to call something and redirect its output to a file, you'll have to call CMD.EXE like this:

lResult = ShellAndWait("cmd.exe /c MyCommand.EXE > MyFile.TMP", 10000)
 
### More Info
 
strCommandLine As String: the command line to execute;

lWait As Long: the number of milliseconds to wait for the process to finish, or 0 to wait indefinately

This code is not original; I am not its author. I believe it was originally available as an MSDN KB article.

Application return code on Success, -1 on Error


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Adam Murray](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/adam-murray.md)
**Level**          |Advanced
**User Rating**    |4.8 (19 globes from 4 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/adam-murray-vb-code-to-shell-and-wait__1-14938/archive/master.zip)

### API Declarations

```
Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As Long, ByVal lpCommandLine As String, ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, ByVal lpEnvironment As Long, ByVal lpCurrentDriectory As Long, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Type PROCESS_INFORMATION
 hProcess As Long
 hThread As Long
 dwProcessId As Long
 dwThreadId As Long
End Type
Private Type STARTUPINFO
 cb As Long
 lpReserved As Long
 lpDesktop As Long
 lpTitle As Long
 dwX As Long
 dwY As Long
 dwXSize As Long
 dwYSize As Long
 dwXCountChars As Long
 dwYCountChars As Long
 dwFillAttribute As Long
 dwFlags As Long
 wShowWindow As Integer
 cbReserved2 As Integer
 lpReserved2 As Long
 hStdInput As Long
 hStdOutput As Long
 hStdError As Long
End Type
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
```


### Source Code

```
Function ShellAndWait(strCommandLine As String, lWait As Long) As Long
 Dim objProcess As PROCESS_INFORMATION
 Dim objStartup As STARTUPINFO
 Dim lResult As Long
 Dim lExitCode As Long
 objStartup.cb = 68
 objStartup.lpReserved = 0
 objStartup.lpDesktop = 0
 objStartup.lpTitle = 0
 objStartup.dwX = 0
 objStartup.dwY = 0
 objStartup.dwXSize = 0
 objStartup.dwYSize = 0
 objStartup.dwXCountChars = 0
 objStartup.dwYCountChars = 0
 objStartup.dwFillAttribute = 0
 objStartup.dwFlags = 0
 objStartup.wShowWindow = 0
 objStartup.cbReserved2 = 0
 objStartup.lpReserved2 = 0
 objStartup.hStdInput = 0
 objStartup.hStdOutput = 0
 objStartup.hStdError = 0
 'try and Create the process
 lResult = CreateProcess(0, strCommandLine, 0, 0, 0, 0, 0, 0, objStartup, objProcess)
 If lResult = 0 Then
 ShellAndWait = -1
 Exit Function
 End If
 'now, wait on the process
 If lWait <> 0 Then
 lResult = WaitForSingleObject(objProcess.hProcess, lWait)
 If lResult = 258 Then 'did we timeout?
 lResult = TerminateProcess(objProcess.hProcess, -1)
 lResult = WaitForSingleObject(objProcess.hProcess, lWait)
 End If
 End If
 'let's get the exit code from the process
 lResult = GetExitCodeProcess(objProcess.hProcess, lExitCode)
 lResult = CloseHandle(objProcess.hProcess)
 lResult = CloseHandle(objProcess.hThread)
 ShellAndWait = lExitCode
End Function
```

