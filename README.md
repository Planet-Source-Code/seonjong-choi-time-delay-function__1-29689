<div align="center">

## Time Delay Function


</div>

### Description

It is method that delay time as much as time that designate.

There is no problem even if use over midnight.

With Sleep function, use and lowered CPU's use rate.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SeonJong,CHOI](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/seonjong-choi.md)
**Level**          |Beginner
**User Rating**    |4.3 (26 globes from 6 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/seonjong-choi-time-delay-function__1-29689/archive/master.zip)

### API Declarations

```
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
```


### Source Code

```
Public Sub Delay(ByVal delayTime As Single)
  Dim startTime  As Single
  Dim LoopTime  As Single
  startTime = Timer
  Do
    Call Sleep(50)
    If Timer < startTime Then startTime = startTime - 86400
    LoopTime = Timer - startTime
    DoEvents
  Loop Until delayTime < LoopTime
End Sub
```

