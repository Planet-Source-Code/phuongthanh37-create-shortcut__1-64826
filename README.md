<div align="center">

## Create shortcut


</div>

### Description

Create Shortcut
 
### More Info
 
Shortcut Name, Target Name, Start In, SIcon...

Shortcut link to target name


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[phuongthanh37](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/phuongthanh37.md)
**Level**          |Beginner
**User Rating**    |5.0 (15 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Files/ File Controls/ Input/ Output](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/files-file-controls-input-output__1-3.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/phuongthanh37-create-shortcut__1-64826/archive/master.zip)





### Source Code

```
Option Explicit
Function TaoShortCut(TenFileShortCut As String, sName As String, _
      Optional sParam As String, Optional sStratIn As String, _
      Optional sIcon As String, Optional sComment As String)
Dim OBJ As Object, oShellLink As Object
Set OBJ = CreateObject("wscript.shell")
Set oShellLink = OBJ.CreateShortcut(TenFileShortCut)
With oShellLink
  .TargetPath = sName
  .Arguments = sParam
  .WorkingDirectory = sStratIn
  If sIcon = "" Then sIcon = sName
  .IconLocation = sIcon
  .Description = sComment
  .Save
End With
End Function
Function TaoShortCutOnDeskTop(TenFileShortCut As String, sName As String, _
      Optional sParam As String, Optional sStratIn As String, _
      Optional sIcon As String, Optional sComment As String)
Dim OBJ As Object, oShellLink As Object
Set OBJ = CreateObject("wscript.shell")
Set oShellLink = OBJ.CreateShortcut(TenFileShortCut)
Set oShellLink = OBJ.CreateShortcut(OBJ.SpecialFolders("Desktop") & "\" & TenFileShortCut)
With oShellLink
  .TargetPath = sName
  .Arguments = sParam
  .WorkingDirectory = sStratIn
  If sIcon = "" Then sIcon = sName
  .IconLocation = sIcon
  .Description = sComment
  .Save
End With
End Function
Private Sub Form_Load()
TaoShortCut "C:\Short1.Lnk", "E:\WINDOWS\system32\notepad.exe", , , , "Create shortcut by phuongthanh37"
TaoShortCutOnDeskTop "Short2.lnk", "%SystemRoot%\explorer.exe", "/e", "%HOMEDRIVE%%HOMEPATH%", "%SystemRoot%\explorer.exe,1", "Create shortcut by phuongthanh37"
End Sub
```

