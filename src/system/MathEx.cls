VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Math"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Description("Provides methods to extend native math operations.")
'@Folder("VBALibrary.System")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")

Option Explicit

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
