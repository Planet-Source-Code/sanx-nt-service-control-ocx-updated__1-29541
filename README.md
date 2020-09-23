<div align="center">

## NT Service Control OCX \*\*UPDATED\*\*


</div>

### Description

**UPDATED** A small OCX that can be used to Stop, Start, Pause, Resume, query the current state and change the start type of any Windows NT Service. Works on NT4, 2000 and XP.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |2001-12-17 11:50:02
**By**             |[Sanx](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/sanx.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[NT\_Service4251412162001\.zip](https://github.com/Planet-Source-Code/sanx-nt-service-control-ocx-updated__1-29541/archive/master.zip)

### API Declarations

```
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal strMachineName As String, ByVal strDBName As String, ByVal lAccessReq As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal strServiceName As String, ByVal lAccessReq As Long) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal lNumServiceArgs As Long, ByVal strArgs As String) As Boolean
Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal lControlCode As Long, lpServiceStatus As SERVICE_STATUS) As Boolean
Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hHandle As Long) As Boolean
Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Boolean
```





