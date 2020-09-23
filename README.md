<div align="center">

## How to detect if mousewheel scrolls


</div>

### Description

This snippet will detect if the mouse wheel scrolls, it does not detect if there is a scroll up/down, i do not know how to do that.
 
### More Info
 
Add a timer, set inteval to 1.

Add a listbox.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ian Gorse](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ian-gorse.md)
**Level**          |Intermediate
**User Rating**    |3.8 (15 globes from 4 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ian-gorse-how-to-detect-if-mousewheel-scrolls__1-25633/archive/master.zip)

### API Declarations

```
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Type POINTAPI
  x As Long
  y As Long
End Type
Private Type Msg
  hwnd As Long
  message As Long
  wParam As Long
  lParam As Long
  time As Long
  pt As POINTAPI
End Type
```


### Source Code

```
on the timer function, type....
Dim amsg As Msg
GetMessage amsg, 0, 0, 0
DispatchMessage amsg
If amsg.message = 522 Then
 list1.additem "Mouse wheel scrolled"
end if
'that is all, hope it comes useful.
```

