<div align="center">

## Color Conversion Functions


</div>

### Description

Convert any color format (hex, long, rgb) to any other color format. There may be other examples of this on PSC, but I checked them and they do not use the same algorithm. This is not the fastest way to convert colors, but it is simple and reliable.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Simon Price](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/simon-price.md)
**Level**          |Beginner
**User Rating**    |4.6 (37 globes from 8 users)
**Compatibility**  |VB 6\.0
**Category**       |[Graphics](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/graphics__1-46.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/simon-price-color-conversion-functions__1-26213/archive/master.zip)





### Source Code

```
' *** COLOR CONVERSTION FUNCTIONS ***
' this is the main function, all the other converstion functions play off this 1
' accepted input hex formats: &H######, ######, #*****
' NOT: &H#***** !!! (i hope no1 would use that anyway)
Public Sub Hex2RGB(strHexColor As String, r As Byte, g As Byte, b As Byte)
Dim HexColor As String
Dim i As Byte
On Error Resume Next
  ' make sure the string is 6 characters long
  ' (it may have been given in &H###### format, we want ######)
  strHexColor = Right((strHexColor), 6)
  ' however, it may also have been given as or #***** format, so add 0's in front
  For i = 1 To (6 - Len(strHexColor))
    HexColor = HexColor & "0"
  Next
  HexColor = HexColor & strHexColor
  ' convert each set of 2 characters into bytes, using vb's cbyte function
  r = CByte("&H" & Right$(HexColor, 2))
  g = CByte("&H" & Mid$(HexColor, 3, 2))
  b = CByte("&H" & Left$(HexColor, 2))
End Sub
Public Function RGB2Hex(r As Byte, g As Byte, b As Byte) As String
On Error Resume Next
  ' convert to long using vb's rgb function, then use the long2rgb function
  RGB2Hex = Long2Hex(RGB(r, g, b))
End Function
Public Sub Long2RGB(LongColor As Long, r As Byte, g As Byte, b As Byte)
On Error Resume Next
  ' convert to hex using vb's hex function, then use the hex2rgb function
  Hex2RGB (Hex(LongColor))
End Sub
Public Function RGB2Long(r As Byte, g As Byte, b As Byte) As Long
On Error Resume Next
  ' use vb's rgb function
  RGB2Long = RGB(r, g, b)
End Function
Public Function Long2Hex(LongColor As Long) As String
On Error Resume Next
  ' use vb's hex function
  Long2Hex = Hex(LongColor)
End Function
Public Function Hex2Long(strHexColor As String) As Long
Dim r As Byte
Dim g As Byte
Dim b As Byte
On Error Resume Next
  ' use the hex2rgb function to get the red green and blue bytes
  Hex2RGB strHexColor, r, g, b
  ' convert to long using vb's rgb function
  Hex2Long = RGB(r, g, b)
End Function
```

