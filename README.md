<div align="center">

## encryptAll


</div>

### Description

This simple encryption Function uses Xor encryption and itterative encoding to more securely encrypt a text string. Although the code is not unbreakable, it cannot be broken with a simple key. Itterative encoding ensures that code-cracking techniques, like character frequency study, will not work. It's as easy as could be. Function can be easily modified to encode more than strings.
 
### More Info
 
data As String, seed As Long Integer

Module is set to Option Base 1. This is to simplify the code. If you are adding this to a pre-existing module, make sure that Base Option 1 will not adversely affect your other functions, or modify the code to work off of Base Option 0.

Function returns String of encoded/decoded text.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Bryan Hurley](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/bryan-hurley.md)
**Level**          |Unknown
**User Rating**    |4.2 (161 globes from 38 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Encryption](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/encryption__1-48.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/bryan-hurley-encryptall__1-2977/archive/master.zip)





### Source Code

```
Public Function encryptAll(data As String, seed As Long) As String
Dim x As Integer, tmp As String, stepnum As Integer
Dim byteArray() As Byte, seedOffset As Integer, n As String
tmp = Trim$(Str(seed))
seed = 0
For x = 1 To Len(tmp)
n = Mid(tmp, x, 1)
 seed = seed + CLng(n)
Next x
reCheckSeed:
 If seed > 255 Then
  seed = -1 + (seed - 255)
  GoTo reCheckSeed
 End If
For x = 1 To Len(data)
 ReDim Preserve byteArray(x)
 n = Mid(data, x, 1)
 byteArray(x) = Asc(n)
 stepnum = seed + x
reCheckstepnum:
 If stepnum > 255 Then
  stepnum = -1 + (stepnum - 255)
  GoTo reCheckstepnum
 End If
 byteArray(x) = byteArray(x) Xor CByte(stepnum)
Next x
 tmp = ""
 For x = 1 To Len(data)
  tmp = tmp & Chr(byteArray(x))
 Next x
encryptAll = tmp
End Function
```

