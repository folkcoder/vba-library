VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MathEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Description("Provides methods to extend native math operations.")
'@Folder("VBALibrary.System")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")

Option Explicit

Private Const mlngTypeMismatchErrorNumber as Long = 13
Private Const mlngErrorNumber As Long = vbObjectError + 513
Private Const mstrClassName As String = "MathEx"


'@Description("Calculates the nearest multiple of the specified factor, rounded up.")
'@Param("dblNumber: The double for which a ceiling is being calculated.")
'@Param("dblFactor: The multiple to be used.")
'@Returns("The nearest multiple of the specified factor, rounded up.")
Public Function Ceiling(ByVal dblNumber As Double, Optional ByVal dblFactor As Double = 1) As Double
    Ceiling = (Int(dblNumber / dblFactor) - (dblNumber / dblFactor - Int(dblNumber / dblFactor) > 0)) * dblFactor
End Function


'@Description("Calculates the nearest multiple of the specified factor, rounded down.")
'@Param("dblNumber: The double for which a floor is being calculated.")
'@Param("dblFactor: The multiple to be used.")
'@Returns("The nearest multiple of the specified factor, rounded down.")
Public Function Floor(ByVal dblNumber As Double, Optional ByVal dblFactor As Double = 1) As Double
    Floor = Int(dblNumber / dblFactor) * dblFactor
End Function


'@Description("Returns the larger of the two provided numbers.")
'@Param("vntNumberValues: An array of all the number to be evaluated.")
'@Returns("The smallest number in the provided array.")
Public Function Max(ParamArray vntNumberValues() As Variant) As Variant

Dim vntMaximumValue As Variant
Dim vntCurrentValue As Variant
Dim lngIndex As Long

    vntMaximumValue = vntNumberValues(0)
    For lngIndex = 1 To UBound(vntNumberValues)
        
        vntCurrentValue = vntNumberValues(lngIndex)
        If Not (IsDate(vntCurrentValue) Or IsNumeric(vntCurrentValue)) Then
            Err.Raise mlngTypeMismatchErrorNumber, "MathEx.Max", "One or more of the parameters is not a numeric data type."
        End If
        
        If vntCurrentValue > vntMaximumValue Then
            vntMaximumValue = vntCurrentValue
        End If
        
    Next lngIndex
    Max = vntMaximumValue
    
End Function


'@Description("Returns the smaller of the two provided numbers.")
'@Param("vntNumberValues: An array of all the number to be evaluated.")
'@Returns("The smallest number in the provided array.")
Public Function Min(ParamArray vntNumberValues() As Variant) As Variant

Dim vntMinimumValue As Variant
Dim vntCurrentValue As Variant
Dim lngIndex As Long
   
    vntMinimumValue = vntNumberValues(0)
    For lngIndex = 1 To UBound(vntNumberValues)
        
        vntCurrentValue = vntNumberValues(lngIndex)
        If Not (IsDate(vntCurrentValue) Or IsNumeric(vntCurrentValue)) Then
            Err.Raise mlngTypeMismatchErrorNumber, "MathEx.Min", "One or more of the parameters is not a numeric data type."
        End If
        
        If vntCurrentValue < vntMinimumValue Then
            vntMinimumValue = vntCurrentValue
        End If
    
    Next lngIndex

    Min = vntMinimumValue
       
End Function
