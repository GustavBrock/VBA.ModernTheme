Attribute VB_Name = "ColourHandling"
' Generic functions for conversion of colour values.
' Supplements the native VBA.RGB function.
' 2022-09-25. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.1.0
' License: MIT.
'
Option Explicit

' Create bitmap file of one pixel having the specified RGB colour.
'
' 2022-09-25. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function CreateBitmapFile( _
    ByVal FileName As String, _
    ByVal ColorR As Byte, _
    ByVal ColorG As Byte, _
    ByVal ColorB As Byte) _
    As Boolean
    
    ' White 1-pixel BMP file. the three FF bytes are the colour information.
    Const BitMapMask    As String = _
        "424D3A000000000000003600000028000000010000000100000001001800000000000400000000000000000000000000000000000000FFFFFF00"
        
    ' Index for the colour information in the bitmap file.
    Const ColorIndexR   As Integer = 56
    Const ColorIndexG   As Integer = 55
    Const ColorIndexB   As Integer = 54
    
    Dim FileLength      As Integer
    Dim FileNumber      As Integer
    Dim Index           As Integer
    Dim Bytes()         As Byte
    Dim Success         As Boolean
    
    FileLength = Len(BitMapMask) / 2
    ReDim Bytes(0 To FileLength - 1)
    
    ' Build file bytes.
    For Index = LBound(Bytes) To UBound(Bytes) Step 1
        Bytes(Index) = CByte("&H" & Mid(BitMapMask, 1 + Index * 2, 2))
    Next
    ' Fill in colours.
    Bytes(ColorIndexR) = ColorR
    Bytes(ColorIndexG) = ColorG
    Bytes(ColorIndexB) = ColorB
    
    ' Write bitmap file.
    FileNumber = FreeFile
    Open FileName For Binary As #FileNumber
    Put #FileNumber, , Bytes()
    Close #FileNumber
    
    Success = CBool(Len(Dir(FileName, vbNormal)))
    
    CreateBitmapFile = Success

End Function

' Calculate discrete RGB colours from a composite colour value and
' return one component.
' Also, by reference, return all components.
'
' Examples:
'   Simple print of the components:
'
'   SomeColor = 813466
'   RGBComponent SomeColor
'   ' Debug Print:
'   ' 154           105           12
'
'   Get one component from a colour value:
'
'   Dim SomeColor   As Long
'   Dim Green       As Integer
'   SomeColor = 13466
'   Green = RGBComponent(SomeColor, vbGreen)
'   ' Green ->  52
'
'   Get all components from a colour value:
'
'   Dim SomeColor   As Long
'   Dim Red         As Integer
'   Dim Green       As Integer
'   Dim Blue        As Integer
'   SomeColor = 813466
'   RGBComponent SomeColor, , Red, Green, Blue
'   ' Red   -> 154
'   ' Green -> 105
'   ' Green ->  12
'
' 2017-03-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RGBComponent( _
    ByVal RGB As Long, _
    Optional ByVal Component As Long, _
    Optional ByRef Red As Integer, _
    Optional ByRef Green As Integer, _
    Optional ByRef Blue As Integer) _
    As Integer
    
    Dim Color   As Long
  
    If RGB <= 0 Then
        ' Return Black.
        Red = 0
        Green = 0
        Blue = 0
    Else
        ' Extract the discrete colours from the composite RGB.
        Red = RGB And vbRed
        Green = (RGB And vbGreen) / &H100
        Blue = (RGB And vbBlue) / &H10000
        ' Return chosen colour component.
        Select Case Component
            Case vbRed
                Color = Red
            Case vbGreen
                Color = Green
            Case vbBlue
                Color = Blue
            Case Else
                Color = vbBlack
        End Select
    End If
    
    Debug.Print Red, Green, Blue
    
    RGBComponent = Color

End Function

' Returns the numeric RGB value from an CSS RGB hex representation.
' Will accept strings with or without a leading octothorpe.
'
' Examples:
'   Color = RGBCompound("#9A690C")
'   ' Color = 813466
'   Color = RGBCompound("9A690C")
'   ' Color = 813466
'
' 2017-03-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RGBCompound( _
    ByVal HexRGB As String) _
    As Long
    
    ' Format of RGB hex strings.
    Const RGBPrefix As String = "#"
    Const Length    As Integer = 6
    ' Format of Hex values.
    Const HexPrefix As String = "&H"
    
    Dim Start       As Integer
    Dim Color       As Long
    
    If Mid(HexRGB, 1, 1) = RGBPrefix Then
        Start = 1
    End If
    If Len(HexRGB) = Start + Length Then
        Color = RGB( _
            HexPrefix & Mid(HexRGB, Start + 1, 2), _
            HexPrefix & Mid(HexRGB, Start + 3, 2), _
            HexPrefix & Mid(HexRGB, Start + 5, 2))
    End If
    
    RGBCompound = Color
    
End Function

' Returns the CSS hex representation of a decimal RGB value
' with or without a leading octothorpe.
'
' Example:
'   CSSValue = RGBHex(813466)
'   ' CSSValue = "#9A690C"
'   CSSValue = RGBHex(813466, True)
'   ' CSSValue = "9A690C"
'
' 2017-03-26. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RGBHex( _
    ByVal Color As Long, _
    Optional ByVal NoPrefix As Boolean) _
    As String
    
    ' Format of RGB hex strings.
    Const RGBPrefix As String = "#"
    
    Dim Red     As Integer
    Dim Green   As Integer
    Dim Blue    As Integer
    Dim HexRGB  As String
    
    RGBComponent Color, , Red, Green, Blue
    
    If Not NoPrefix Then
        ' Set prefix.
        HexRGB = RGBPrefix
    End If
    ' Assemble compound string with leading zeroes for small values.
    HexRGB = HexRGB & _
        Right("0" & Hex(Red), 2) & _
        Right("0" & Hex(Green), 2) & _
        Right("0" & Hex(Blue), 2)
    
    RGBHex = HexRGB
    
End Function

