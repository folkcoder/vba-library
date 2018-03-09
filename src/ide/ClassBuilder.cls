VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClassBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.IDE")
'@Descriptions("Provides methods to quickly create a VBA class object stub containing boilerplate property accessors with private backing fields.")
Option Explicit

Private Const mstrClassName As String = "ClassBuilder"
Private Const mlngErrorNumber As Long = vbObjectError + 513
Private Const mstrHeaderSeparator As String = "' ============================================================================="

Public Enum AccessLevel
    FriendAccess
    PrivateAccess
    PublicAccess
End Enum

Public Enum DataTypeOption
    BooleanType
    ByteType
    CollectionType
    CurrencyType
    DateType
    DoubleType
    IntegerType
    LongType
    ObjectType
    SingleType
    StringType
    VariantType
End Enum

Private Type TClassBuilder
    ClassName As String
    Fields As Object
    IsInitialized As Boolean
    Properties As Object
End Type

Private this As TClassBuilder


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    Set this.Fields = CreateObject("System.Collections.Sortedlist")
    Set this.Properties = CreateObject("System.Collections.Sortedlist")
End Sub

'@Description("Initializes the ClassBuilder instance.")
'@Param("strClassName: The name of the class to be created.")
Public Sub Initialize(ByVal strClassName As String)
    this.ClassName = strClassName
    this.IsInitialized = True
End Sub


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Adds property and backing field boilerplate to the ClassBuilder instance.")
'@Param("strPropertyName: The name of the property.")
'@Param("PropertyDataTypeOption: The property's data type.")
'@Param("blnCreateBackingField: A value indicating whether the member should include a private backing field.")
'@Param("blnCreateGetProperty: A value indicating whether the member should include a get accessor.")
'@Param("GetPropertyAccessLevel: The access level of the get accessor.")
'@Param("blnCreateLetProperty: A value indicating whether the member should include a let accessor.")
'@Param("LetPropertyAccessLevel: The access level of the let accessor.")
Public Sub AddMember( _
    ByVal strPropertyName As String, _
    ByVal PropertyDataTypeOption As DataTypeOption, _
    Optional ByVal blnCreateBackingField As Boolean = True, _
    Optional ByVal blnCreateGetProperty As Boolean = True, _
    Optional ByVal GetPropertyAccessLevel As AccessLevel = AccessLevel.PublicAccess, _
    Optional ByVal blnCreateLetProperty As Boolean = True, _
    Optional ByVal LetPropertyAccessLevel As AccessLevel = AccessLevel.PublicAccess _
    )
 
Dim strFieldText As String
Dim strPropertyText As String

Dim strDataTypeOptionName As String
Dim blnIsObject As Boolean

    VerifyInitialization

    strDataTypeOptionName = GetDataTypeOptionName(PropertyDataTypeOption)
    blnIsObject = IsObject(PropertyDataTypeOption)
    
    ' Construct field
    If blnCreateBackingField Then
        strFieldText = strPropertyName & " As " & strDataTypeOptionName
        this.Fields.Add strPropertyName, strFieldText
    End If
    
    ' Construct get accessor
    If blnCreateGetProperty Then
        strPropertyText = strPropertyText & GetAccessLevelName(GetPropertyAccessLevel) & " Property Get " & strPropertyName & "() As " & strDataTypeOptionName & vbNewLine
        If blnCreateBackingField Then
            strPropertyText = strPropertyText & vbTab
            If blnIsObject Then
                strPropertyText = strPropertyText & "Set "
            End If
            strPropertyText = strPropertyText & strPropertyName & " = " & "this." & strPropertyName & vbNewLine
        End If
        strPropertyText = strPropertyText & "End Property"
    End If
    
    ' Construct let accessor
    If blnCreateLetProperty Then
        
        ' handle spacing if property includes a get accessor
        If Not strPropertyText = vbNullString Then
            strPropertyText = strPropertyText & vbNewLine
        End If
        
        strPropertyText = strPropertyText & GetAccessLevelName(LetPropertyAccessLevel) & " Property Let " & strPropertyName & "(ByVal Value As " & strDataTypeOptionName & ")" & vbNewLine
        If blnCreateBackingField Then
            strPropertyText = strPropertyText & vbTab
            If blnIsObject Then
                strPropertyText = strPropertyText & "Set "
            End If
            strPropertyText = strPropertyText & "this." & strPropertyName & " = Value" & vbNewLine
        End If
        strPropertyText = strPropertyText & "End Property"
    End If
    
    If Not strPropertyText = vbNullString Then
        this.Properties.Add strPropertyName, strPropertyText
    End If

End Sub

' =============================================================================

'@Description("Generates a VBA Class Module from the current ClassBuilder instance.")
Public Sub GenerateClass()

Const vbext_ct_ClassModule = 2

Dim objVbaEditor As Object
Dim objVbaProject As Object
Dim objVbaComponent As Object
Dim objCodeModule As Object
Dim lngLineNumber As Long

Dim intFieldsCounter As Integer
Dim intPropertiesCounter As Integer

    VerifyInitialization

    Set objVbaEditor = Application.VBE
    Set objVbaProject = objVbaEditor.ActiveVBProject
    Set objVbaComponent = objVbaProject.VBComponents.Add(vbext_ct_ClassModule)
    Set objCodeModule = objVbaComponent.CodeModule
    
    objVbaComponent.Name = mstrClassName
    With objCodeModule
        
        lngLineNumber = 1
        
        ' Declaration
        Do Until .CountOfDeclarationLines = 0
            .DeleteLines lngLineNumber
        Loop
        .InsertLines 1, "Option Explicit" & vbNewLine
        lngLineNumber = .CountOfLines + 1
        
        ' Fields
        .InsertLines lngLineNumber, "Private Type T" & mstrClassName
        lngLineNumber = .CountOfLines + 1
        For intFieldsCounter = 0 To this.Fields.Count - 1
            .InsertLines lngLineNumber, vbTab & " " & this.Fields.GetByIndex(intFieldsCounter)
            lngLineNumber = .CountOfLines + 1
        Next intFieldsCounter
        .InsertLines lngLineNumber, "End Type"
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, vbNewLine
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, "Private this As T" & mstrClassName
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, vbNewLine
        lngLineNumber = .CountOfLines + 1
        
        ' Constructor
        .InsertLines lngLineNumber, CreateConstructorRegion
        
        ' Properties
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, CreatePropertiesHeader
        For intPropertiesCounter = 0 To this.Properties.Count - 1
            lngLineNumber = .CountOfLines + 1
            .InsertLines lngLineNumber, vbNewLine & this.Properties.GetByIndex(intPropertiesCounter)
        Next intPropertiesCounter
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, vbNewLine

        ' Methods
        lngLineNumber = .CountOfLines + 1
        .InsertLines lngLineNumber, CreateMethodsHeader
        
    End With
    
End Sub


' =============================================================================
' PRIVATE METHODS
' =============================================================================

'@Description("Constructs a header for the class's constructor/destructor.")
'@Returns("Boilerplate header for the class's constructor/destructor.")
Private Function CreateConstructorRegion() As String
    CreateConstructorRegion = vbNullString & _
        mstrHeaderSeparator & vbNewLine & _
        "' CONSTRUCTOR / DESTRUCTOR" & vbNewLine & _
        mstrHeaderSeparator & vbNewLine & vbNewLine
End Function

' =============================================================================

'@Description("Constructs a header for the class's methods section.")
'@Returns("Boilerplate header for the class's methods section.")
Private Function CreateMethodsHeader() As String
    CreateMethodsHeader = vbNullString & _
        mstrHeaderSeparator & vbNewLine & _
        "' METHODS" & vbNewLine & _
        mstrHeaderSeparator & vbNewLine
End Function

' =============================================================================

'@Description("Constructs a header for the class's properties section.")
'@Returns("Boilerplate header for the class's properties section.")
Private Function CreatePropertiesHeader() As String
    CreatePropertiesHeader = vbNullString & _
        mstrHeaderSeparator & vbNewLine & _
        "' PROPERTIES" & vbNewLine & _
        mstrHeaderSeparator
End Function

' =============================================================================

'@Description("Converts a AccessLevel enumerated value into its string representation.")
'@Param("AccessLevel: The AccessLevel to be converted.")
'@Returns("A string representation of the provided AccessLevel value.")
Private Function GetAccessLevelName(ByVal AccessLevelEnum As AccessLevel) As String
    
    Select Case AccessLevelEnum
        Case AccessLevel.FriendAccess:
            GetAccessLevelName = "Friend"
        Case AccessLevel.PrivateAccess:
            GetAccessLevelName = "Private"
        Case AccessLevel.PublicAccess:
            GetAccessLevelName = "Public"
    End Select

End Function

' =========================================================================

'@Description("Converts a DataTypeOption enumerated value into its string representation.")
'@Param("DataTypeOptionEnum: The DataTypeOption to be converted into a string.")
'@Returns("A string representation of the provided DataTypeOption value.")
Private Function GetDataTypeOptionName(ByVal DataTypeOptionEnum As DataTypeOption) As String

    Select Case DataTypeOptionEnum
        Case DataTypeOption.BooleanType:
            GetDataTypeOptionName = "Boolean"
        Case DataTypeOption.ByteType:
            GetDataTypeOptionName = "Byte"
        Case DataTypeOption.CollectionType:
            GetDataTypeOptionName = "Collection"
        Case DataTypeOption.CurrencyType:
            GetDataTypeOptionName = "Currency"
        Case DataTypeOption.DateType:
            GetDataTypeOptionName = "Date"
        Case DataTypeOption.DoubleType:
            GetDataTypeOptionName = "Double"
        Case DataTypeOption.IntegerType:
            GetDataTypeOptionName = "Integer"
        Case DataTypeOption.LongType:
            GetDataTypeOptionName = "Long"
        Case DataTypeOption.ObjectType:
            GetDataTypeOptionName = "Object"
        Case DataTypeOption.SingleType:
            GetDataTypeOptionName = "Single"
        Case DataTypeOption.StringType:
            GetDataTypeOptionName = "String"
        Case DataTypeOption.VariantType:
            GetDataTypeOptionName = "Variant"
    End Select
    
End Function

' =========================================================================

'@Description("Determines whether the provided DataTypeOption is an object.")
'@Param("DataTypeOptionEnum: The DataTypeOption to be tested.")
'@Returns("A value indicating whether the provided DataTypeOption is an object.")
Private Function IsObject(ByVal DataTypeOptionEnum As DataTypeOption) As Boolean
    Select Case DataTypeOptionEnum
        Case DataTypeOption.CollectionType, DataTypeOption.ObjectType:
            IsObject = True
        Case Else
            IsObject = False
    End Select
End Function

' =========================================================================

'@Description("Verifies that the ClassBuilder instance has been properly initialized").
Private Sub VerifyInitialization()
    If Not this.IsInitialized Then
        Err.Raise mlngErrorNumber, mstrClassName & ".VerifyInitialization", "The ClassBuilder instance has must be initialized via the Initialize method before it can be used. Consider delegating responsibility for creating the class to a factory method."
    End If
End Sub

' =========================================================================
