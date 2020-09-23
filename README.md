<div align="center">

## Attaching Help to Apps


</div>

### Description

A shell command could be used to do this as *.chm file excute by clicking on them. Below is the shell command to use. Also shows how to shell execute a string variable.

More info on Help and what you could with it. Goto. http://www.smountain.com/c_VBHelp.htm.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[TonyGG](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/tonygg.md)
**Level**          |Intermediate
**User Rating**    |4.8 (29 globes from 6 users)
**Compatibility**  |VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/tonygg-attaching-help-to-apps__1-33349/archive/master.zip)





### Source Code

```
http://www.smountain.com/c_VBHelp.htm
Option Explicit
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub CmdHelp_Click()
Dim Stg1 As String
'shell the help file
Stg1 = App.Path & "\" & "mshelp.chm" 'exchange with your help file name
ShellExecute hwnd, "open", Stg1, "", "", vbNormalFocus
End Sub
```

