VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IMappableMock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBALibrary.Tests.Data.ObjectRelationalMapping")
Option Explicit

Implements IMappable

Private Type TClass
    TestingID As Long
    TestColor As String
    TestName As String
    TestNumber As Long
End Type

Private this As TClass

Private mobjClassProperties As dictionary


' =============================================================================
' PROPERTIES
' =============================================================================

Public Property Get TestingID() As Long
    TestingID = this.TestingID
End Property
Public Property Let TestingID(ByVal Value As Long)
    this.TestingID = Value
End Property

Public Property Get TestColor() As String
    TestColor = this.TestColor
End Property
Public Property Let TestColor(ByVal Value As String)
    this.TestColor = Value
End Property

Public Property Get TestName() As String
    TestName = this.TestName
End Property
Public Property Let TestName(ByVal Value As String)
    this.TestName = Value
End Property

Public Property Get TestNumber() As Long
    TestNumber = this.TestNumber
End Property
Public Property Let TestNumber(ByVal Value As Long)
    this.TestNumber = Value
End Property


' =============================================================================
' INTERFACE IMPLEMENTATIONS
' =============================================================================

Private Property Get IMappable_MappedProperties() As dictionary
    If mobjClassProperties Is Nothing Then
        Set mobjClassProperties = New dictionary
        With mobjClassProperties
            .Add "TestingID", "ID"
            .Add "TestColor", "TestColor"
            .Add "TestName", "TestName"
            .Add "TestNumber", "Test Number"
        End With
    End If
    Set IMappable_MappedProperties = mobjClassProperties
End Property

Private Property Get IMappable_TableName() As String
    IMappable_TableName = "TestClass"
End Property

Private Function IMappable_CreateNew() As IMappable
    Set IMappable_CreateNew = New IMappableMock
End Function
 
Private Function IMappable_GetPropertyValue(ByVal strPropertyName As String) As Variant
    IMappable_GetPropertyValue = CallByName(Me, strPropertyName, VbGet)
End Function

Private Sub IMappable_LetPropertyValue(ByVal strPropertyName As String, ByVal vntValue As Variant)
    CallByName Me, strPropertyName, VbLet, vntValue
End Sub
