<div align="center">

## The Better SendKeys Statement


</div>

### Description

The SendKeys statement is limited only to some keyboard presses...

While this code gives you more...

Such as:

1. You cannot send Space Keystrokes in SendKeys (I tried to do it, or maybe I just don't know how), while you can in HitKey...

HitKey Asc(" ") or

HitKey KeyCodeConstants.vbKeySpace

2. HitKey can also perform CapsLock, ScrollLock, NumLock, or other Lock, and other Keys you can find in your keyboard.. Remember to use VB Constants in "KeyCodeConstants"

code by: Ronald Borla
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Ronald Borla](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/ronald-borla.md)
**Level**          |Beginner
**User Rating**    |5.0 (20 globes from 4 users)
**Compatibility**  |VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/ronald-borla-the-better-sendkeys-statement__1-67411/archive/master.zip)





### Source Code

```
Option Explicit
Private Declare Sub keybd_event Lib "user32" _
  (ByVal bVk As Byte, ByVal bScan As Byte, _
  ByVal dwFlags As Long, ByVal dwExtraInfo As _
  Long)
Private Declare Function MapVirtualKey _
  Lib "user32" Alias "MapVirtualKeyA" _
  (ByVal wCode As Long, ByVal _
  wMapType As Long) As Long
Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Public Sub HitKey(ByVal KeyCode As Integer)
'Send the key down
Call keybd_event(KeyCode, _
   MapVirtualKey(KeyCode, 0), _
   KEYEVENTF_EXTENDEDKEY Or 0, 0)
'Send the key up
Call keybd_event(KeyCode, _
   MapVirtualKey(KeyCode, 0), _
   KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
DoEvents
End Sub
'###############################################
'Example for the use of the code:
'In you Timer Event
Private Sub tmrKeyPress_Timer()
'Sends Key Spaces every timer event
HitKey KeyCodeConstants.vbKeySpace
End Sub
'Use the VB Default KeyCodeConstants for easier
'use for this code.
'###############################################
```

