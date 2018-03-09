VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ColorTranslator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Description("Provides methods to translate among various color formats.")
'@Folder("VBALibrary.Drawing")

Private Const mlngErrorNumber As Long = vbObjectError + 513
Private Const mstrClassName As String = "ColorTranslator"

Option Explicit

' =============================================================================

'@Description("Converts a web hex value (e.g., #FFFFFF) to RGB.")
'@Param("strHexColor: The hex color to convert.")
'@Returns("An RGB color value.")
Public Function WebHexToRGB(ByVal strHexColor As String) As Long

Dim strFormattedHexColor As String
Dim intRed As Integer
Dim intGreen As Integer
Dim intBlue As Integer

    strFormattedHexColor = Replace$(strHexColor, "#", vbNullString)
    If Len(strFormattedHexColor) <> 6 Then
        Err.Raise mlngErrorNumber, mstrClassName & ".WebHexToRGB", "Hex color is not six characters long, exclusive of # sign."
    End If
    
    intRed = Val("&H" & Mid$(strFormattedHexColor, 1, 2))
    intGreen = Val("&H" & Mid$(strFormattedHexColor, 3, 2))
    intBlue = Val("&H" & Mid$(strFormattedHexColor, 5, 2))

    WebHexToRGB = RGB(intRed, intGreen, intBlue)

End Function

' =============================================================================

'@Description("Converts an RGB color value to a hexademical.")
'@Param("lngRGBColor: The RGB color value to convert.")
'@Returns("A hexadecimal color value.")
Public Function RGBToHex(ByVal lngRGBColor As Long) As String
    RGBToHex = Hex$(lngRGBColor)
End Function

' =============================================================================
