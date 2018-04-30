Attribute VB_Name = "ColourHandling"
Option Explicit

' Generic functions for conversion of colour values.
' Supplements the native VBA.RGB function.
' 2018-04-30. Gustav Brock, Cactus Data ApS, CPH.
' Version 1.0.1
' License: MIT.

' *

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
    
    ' Debug.Print Red, Green, Blue
    
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

' Returns the compound RGB value from discrete CMYK values.
' CMYK values must represent integer percent values: 0 to 100.
'
' Examples:
'   Color = RGBCMYK(0, 100, 100, 0)
'   ' Color = 255
'   Color = RGBCMYK(0, 100, 100, 50)
'   ' Color = 128
'   Color = RGBCMYK(0, 0, 0, 100)
'   ' Color = 0
'   Color = RGBCMYK(100, 100, 100, 0)
'   ' Color = 0
'   Color = RGBCMYK(100, 100, 100, 50)
'   ' Color = 0
'   Color = RGBCMYK(0, 0, 0, 0)
'   ' Color = 16777215
'
' 2018-04-30. Gustav Brock, Cactus Data ApS, CPH.
'
Public Function RGBCMYK( _
    ByVal Cyan As Double, _
    ByVal Magenta As Double, _
    ByVal Yellow As Double, _
    ByVal Black As Double) _
    As Long

    ' Minimum and maximum values.
    Const MaxRGB    As Double = &HFF
    Const MaxCMYK   As Double = 100
    Const MinCMYK   As Double = 0
    
    Const Half      As Double = 0.5
    
    Dim Brightness  As Double
    Dim Red         As Integer
    Dim Green       As Integer
    Dim Blue        As Integer
    Dim Color       As Long
    
    ' Limit input to acceptable range for CMYK values.
    If Cyan < MinCMYK Then
        Cyan = MinCMYK
    ElseIf Cyan > MaxCMYK Then
        Cyan = MaxCMYK
    End If
    If Magenta < MinCMYK Then
        Magenta = MinCMYK
    ElseIf Magenta > MaxCMYK Then
        Magenta = MaxCMYK
    End If
    If Yellow < MinCMYK Then
        Yellow = MinCMYK
    ElseIf Yellow > MaxCMYK Then
        Yellow = MaxCMYK
    End If
    If Black < MinCMYK Then
        Black = MinCMYK
    ElseIf Black > MaxCMYK Then
        Black = MaxCMYK
    End If
    
    ' Calculate brightness factor.
    Brightness = Int(MaxRGB / MaxCMYK * (MaxCMYK - Black) + Half) / MaxCMYK
    ' Calculate RGB colours.
    Red = Brightness * (MaxCMYK - Cyan)
    Green = Brightness * MaxCMYK * (MaxCMYK - Magenta)
    Blue = Brightness * MaxCMYK * (MaxCMYK - Yellow)
    
    ' Calculate RGB compound value.
    Color = RGB(Red, Green, Blue)
    
    RGBCMYK = Color
    
End Function

