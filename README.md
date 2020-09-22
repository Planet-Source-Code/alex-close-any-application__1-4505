<div align="center">

## Close Any Application


</div>

### Description

This code will Close any application based on its windows caption
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Alex](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/alex.md)
**Level**          |Unknown
**User Rating**    |4.4 (31 globes from 7 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/alex-close-any-application__1-4505/archive/master.zip)

### API Declarations

```
Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_CLOSE = &H10
```


### Source Code

```

Public Function CloseApplication(byVal sAppCaption As String) As Boolean
  Dim lHwnd As Long
  Dim lRetVal As Long
  lHwnd = FindWindow(vbNullString, sAppCaption)
  If lHwnd <> 0 Then
    lRetVal = PostMessage(lHwnd, WM_CLOSE, 0&, 0&)
  End If
End Function
```

