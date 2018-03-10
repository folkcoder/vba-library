VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMappable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Description("Provides an interface to allow a class to be mapped by an object-relational mapper.")
'@Folder("VBALibrary.Data.ObjectRelationalMapping")
'@Interface

Option Explicit

Public Property Get MappedProperties() As dictionary
Public Property Get MappedProperties() As Dictionary
End Property

Public Property Get TableName() As String
End Property

Public Function CreateNew() As IMappable
End Function

Public Function GetPropertyValue(ByVal strPropertyName As String) As Variant
End Function

Public Sub LetPropertyValue(ByVal strPropertyName As String, ByVal vntValue As Variant)
End Sub
