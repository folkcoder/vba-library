VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MockMappable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("VBALibrary.Tests.Data.ObjectRelationalMapping")
Option Explicit

Implements IMappable

Private Type TClass
    PersonId As Long
    FavoriteColor As String
    PersonName As String
    PersonBirthdate As Date
End Type

Private this As TClass

Private mobjClassProperties As Dictionary


' =============================================================================
' PROPERTIES
' =============================================================================

Public Property Get PersonId() As Long
    PersonId = this.PersonId
End Property
Public Property Let PersonId(ByVal Value As Long)
    this.PersonId = Value
End Property

Public Property Get FavoriteColor() As String
    FavoriteColor = this.FavoriteColor
End Property
Public Property Let FavoriteColor(ByVal Value As String)
    this.FavoriteColor = Value
End Property

Public Property Get PersonName() As String
    PersonName = this.PersonName
End Property
Public Property Let PersonName(ByVal Value As String)
    this.PersonName = Value
End Property

Public Property Get PersonBirthdate() As Date
    PersonBirthdate = this.PersonBirthdate
End Property
Public Property Let PersonBirthdate(ByVal Value As Date)
    this.PersonBirthdate = Value
End Property


' =============================================================================
' INTERFACE IMPLEMENTATIONS
' =============================================================================

Private Property Get IMappable_MappedProperties() As Dictionary
    If mobjClassProperties Is Nothing Then
        Set mobjClassProperties = New Dictionary
        With mobjClassProperties
            .Add "PersonId", "Id"
            .Add "FavoriteColor", "FavoriteColor"
            .Add "PersonName", "Name"
            .Add "PersonBirthdate", "Birth Date"
        End With
    End If
    Set IMappable_MappedProperties = mobjClassProperties
End Property

Private Property Get IMappable_TableName() As String
    IMappable_TableName = "MockMappable"
    ' IMappable_TableName = "MockMappableWithAutonumber"
End Property

Private Function IMappable_CreateNew() As IMappable
    Set IMappable_CreateNew = New MockMappable
End Function
 
Private Function IMappable_GetPropertyValue(ByVal strPropertyName As String) As Variant
    IMappable_GetPropertyValue = CallByName(Me, strPropertyName, VbGet)
End Function

Private Sub IMappable_LetPropertyValue(ByVal strPropertyName As String, ByVal vntValue As Variant)
    CallByName Me, strPropertyName, VbCallType.VbLet, vntValue
End Sub
