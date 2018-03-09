VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Registry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.System.Registry")
'@Description("Provides methods to manage the Windows registry.")
Option Explicit

Public Enum RegistryHive
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Public Enum RegistryProvider
    Bit32 = 32
    Bit64 = 64
End Enum

Private Type TRegistryProvider
    Context As Object
    Registry As Object
    RegistryHive As RegistryHive
End Type

Private this As TRegistryProvider


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()

Dim objLocator As Object

    Set this.Context = CreateObject("WbemScripting.SWbemNamedValueSet")
    this.Context.Add "__ProviderArchitecture", RegistryProvider.Bit64
    
    Set objLocator = CreateObject("Wbemscripting.SWbemLocator")
    Set this.Registry = objLocator.ConnectServer(".", "root\default", vbNullString, vbNullString, , , , this.Context).Get("StdRegProv")

End Sub

' =============================================================================

Private Sub Class_Terminate()
    If Not this.Context Is Nothing Then Set this.Context = Nothing
    If Not this.Registry Is Nothing Then Set this.Registry = Nothing
End Sub


' =============================================================================
' PROPERTIES
' =============================================================================

'@Description("Gets or sets the current registry hive.")
Public Property Get Hive() As RegistryHive
    Hive = this.RegistryHive
End Property
Public Property Let Hive(ByRef Value As RegistryHive)
    this.RegistryHive = Value
End Property

'@Description("Gets or sets the current registry provider.")
Public Property Get Provider() As RegistryProvider
    Provider = this.Context.Item("__ProviderArchitecture")
End Property
Public Property Let Provider(ByRef Value As RegistryProvider)
    this.Context.Item("__ProviderArchitecture") = Value
End Property


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Creates a registry key in the current hive using the current registry provider.")
'@Param("strKeyPath: The key to be created.")
Public Sub CreateRegistryKey(ByVal strKeyPath As String)

Dim objInParameters As Object
Dim objOutParameters As Object

    Set objInParameters = this.Registry.Methods_("CreateKey").InParameters
    objInParameters.hDefKey = this.RegistryHive
    objInParameters.sSubKeyName = strKeyPath
    
    Set objOutParameters = this.Registry.ExecMethod_("CreateKey", objInParameters, , this.Context)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.CreateRegistryKey", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If
   
End Sub

' =============================================================================

'@Description("Recursively deletes a registry tree starting at the provided key path.")
'@Param("strKeyPath: The key to be deleted.")
Public Sub DeleteRegistryTree(ByVal strKeyPath As String)
    DeleteSubkeys strKeyPath
End Sub

' =============================================================================

'@Description("Writes a string value to the provided registry key.")
'@Param("strKeyPath: The key path where the value is located.")
'@Param("strValueName: The name of the value to be modified.")
'@Param("strValue: The value's new data.")
Public Sub WriteRegistryString(ByVal strKeyPath As String, ByVal strValueName As String, ByVal strValue As String)

Dim objInParameters As Object
Dim objOutParameters As Object
      
    Set objInParameters = this.Registry.Methods_("SetStringValue").InParameters
    objInParameters.hDefKey = this.RegistryHive
    objInParameters.sSubKeyName = strKeyPath
    objInParameters.sValueName = strValueName
    objInParameters.sValue = strValue
    
    Set objOutParameters = this.Registry.ExecMethod_("SetStringValue", objInParameters, , this.Context)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.WriteRegistryString", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub

' =============================================================================

'@Description("Writes an integer value to the provided registry key.")
'@Param("strKeyPath: The key path where the value is located.")
'@Param("strValueName: The name of the value to be modified.")
'@Param("intValue: The value's new data.")
Public Sub WriteRegistryInteger(ByVal strKeyPath As String, ByVal strValueName As String, ByVal intValue As Integer)

Dim objInParameters As Object
Dim objOutParameters As Object
   
    Set objInParameters = this.Registry.Methods_("SetDWORDValue").InParameters
    objInParameters.hDefKey = this.RegistryHive
    objInParameters.sSubKeyName = strKeyPath
    objInParameters.sValueName = strValueName
    objInParameters.uValue = intValue
    
    Set objOutParameters = this.Registry.ExecMethod_("SetDWORDValue", objInParameters, , this.Context)
    If Not objOutParameters.ReturnValue = 0 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.WriteRegistryString", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub


' =============================================================================
' PRIVATE METHODS
' =============================================================================

'@Description("Deletes the provided key and all subkeys.")
'@Param("strKeyPath: The key path to be deleted.")
Private Static Sub DeleteSubkeys(ByVal strKeyPath As String)
    
Dim vntSubkeys As Variant
Dim vntSubkey As Variant
    
    this.Registry.EnumKey this.RegistryHive, strKeyPath, vntSubkeys
    If IsArray(vntSubkeys) Then
        For Each vntSubkey In vntSubkeys
            DeleteSubkeys strKeyPath & "\" & CStr(vntSubkey)
        Next vntSubkey
    End If
     
    DeleteRegistryKey strKeyPath
     
End Sub

' =============================================================================

'@Description("Deletes the provided registry key.")
'@Param("strKeyPath: The key path to be deleted.")
Private Static Sub DeleteRegistryKey(ByVal strKeyPath As String)
    
Dim objInParameters As Object
Dim objOutParameters As Object

    Set objInParameters = this.Registry.Methods_("DeleteKey").InParameters
    objInParameters.hDefKey = this.RegistryHive
    objInParameters.sSubKeyName = strKeyPath
    
    Set objOutParameters = this.Registry.ExecMethod_("DeleteKey", objInParameters, , this.Context)
    If Not objOutParameters.ReturnValue = 0 And Not objOutParameters.ReturnValue = 2 Then
        Err.Raise objOutParameters.ReturnValue, "Registry.DeleteRegistryKey", "Error in StdRegProv. WbemErrorEnum " & objOutParameters.ReturnValue & "."
    End If

End Sub

' =============================================================================
