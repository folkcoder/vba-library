VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "COMLibraryRegistration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.System.Registry")
'@Description("Provides methods to manage registration of a COM library.")
Option Explicit

Private Const mstrSoftwareClassesKeyBase32Bit As String = "Software\Classes\WOW6432Node\"
Private Const mstrSoftwareClassesKeyBase64Bit As String = "Software\Classes\"

Private Type TCOMLibraryRegistration
    AssemblyCodebaseFile As String
    AssemblyCodebasePath As String
    AssemblyName As String
    AssemblyVersion As String
    RuntimeVersion As String
    TypeLibraryFileName As String
    TypeLibraryFilePath As String
    TypeLibraryGuid As String
    TypeLibraryName As String
    TypeLibraryVersion As String
End Type

Private mobjRegistry As Registry
Private this As TCOMLibraryRegistration


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    Set mobjRegistry = New Registry
    Me.RegistryHive = HKEY_CURRENT_USER
    Me.RegistryProvider = Bit64
End Sub

Private Sub Class_Terminate()
    If Not mobjRegistry Is Nothing Then Set mobjRegistry = Nothing
End Sub


' =============================================================================
' PROPERTIES
' =============================================================================

'@Description("Gets or sets the codebase file name.")
Public Property Get AssemblyCodebaseFile() As String
    AssemblyCodebaseFile = this.AssemblyCodebaseFile
End Property
Public Property Let AssemblyCodebaseFile(ByVal Value As String)
    this.AssemblyCodebaseFile = Value
End Property

'@Description("Gets or sets the codebase file path.")
Public Property Get AssemblyCodebasePath() As String
    AssemblyCodebasePath = this.AssemblyCodebasePath
End Property
Public Property Let AssemblyCodebasePath(ByVal Value As String)
    this.AssemblyCodebasePath = Value
End Property

'@Description("Gets or sets the assembly's name.")
Public Property Get AssemblyName() As String
    AssemblyName = this.AssemblyName
End Property
Public Property Let AssemblyName(ByVal Value As String)
    this.AssemblyName = Value
End Property

'@Description("Gets or sets the assembly's runtime version.")
Public Property Get RuntimeVersion() As String
    RuntimeVersion = this.RuntimeVersion
End Property
Public Property Let RuntimeVersion(ByVal Value As String)
    this.RuntimeVersion = Value
End Property

'@Description("Gets or sets the assembly's version number.")
Public Property Get AssemblyVersion() As String
    AssemblyVersion = this.AssemblyVersion
End Property
Public Property Let AssemblyVersion(ByVal Value As String)
    this.AssemblyVersion = Value
End Property

'@Description("Gets or sets the registry hive within the COM component should be registered.")
Public Property Get RegistryHive() As RegistryHive
    RegistryHive = mobjRegistry.Hive
End Property
Public Property Let RegistryHive(ByVal Value As RegistryHive)
    mobjRegistry.Hive = Value
End Property

'@Description("Gets or sets the registry provider to which the COM component should be registered.")
Public Property Get RegistryProvider() As RegistryProvider
    RegistryProvider = mobjRegistry.Provider
End Property
Public Property Let RegistryProvider(ByVal Value As RegistryProvider)
    mobjRegistry.Provider = Value
End Property

'@Description("Gets or sets the type library's file name.")
Public Property Get TypeLibraryFileName() As String
    TypeLibraryFileName = this.TypeLibraryFileName
End Property
Public Property Let TypeLibraryFileName(ByVal Value As String)
    this.TypeLibraryFileName = Value
End Property

'@Description("Gets or sets the type library's file path.")
Public Property Get TypeLibraryFilePath() As String
    TypeLibraryFilePath = this.TypeLibraryFilePath
End Property
Public Property Let TypeLibraryFilePath(ByVal Value As String)
    this.TypeLibraryFilePath = Value
End Property

'@Description("Gets or sets the type library's GUID.")
Public Property Get TypeLibraryGuid() As String
    TypeLibraryGuid = this.TypeLibraryGuid
End Property
Public Property Let TypeLibraryGuid(ByVal Value As String)
    this.TypeLibraryGuid = Value
End Property

'@Description("Gets or sets the type library's name.")
Public Property Get TypeLibraryName() As String
    TypeLibraryName = this.TypeLibraryName
End Property
Public Property Let TypeLibraryName(ByVal Value As String)
    this.TypeLibraryName = Value
End Property

'@Description("Gets or sets the type library's version.")
Public Property Get TypeLibraryVersion() As String
    TypeLibraryVersion = this.TypeLibraryVersion
End Property
Public Property Let TypeLibraryVersion(ByVal Value As String)
    this.TypeLibraryVersion = Value
End Property


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Deregisters a COM class.")
'@Param("strGuid: The GUID of the class to be deregistered.")
'@Param("strProgId: The ProgID of the class to be deregistered.")
Public Static Sub DeregisterClass(ByVal strGuid As String, ByVal strProgId As String)

Dim strClassRegistrationKeyBase32Bit As String
Dim strClsidKey32Bit As String

Dim strClassRegistrationKeyBase64Bit As String
Dim strClsidKey64Bit As String

    ' SOFTWARE\CLASSES\ProgId
    strClassRegistrationKeyBase32Bit = mstrSoftwareClassesKeyBase32Bit & strProgId
    mobjRegistry.DeleteRegistryTree strClassRegistrationKeyBase32Bit
               
    strClassRegistrationKeyBase64Bit = mstrSoftwareClassesKeyBase64Bit & strProgId
    mobjRegistry.DeleteRegistryTree strClassRegistrationKeyBase64Bit
               
               
    ' SOFTWARE\CLASSES\CLSID\ClassGuid
    strClsidKey32Bit = mstrSoftwareClassesKeyBase32Bit & "CLSID\" & strGuid
    mobjRegistry.DeleteRegistryTree strClsidKey32Bit
        
    strClsidKey64Bit = mstrSoftwareClassesKeyBase64Bit & "CLSID\" & strGuid
    mobjRegistry.DeleteRegistryTree strClsidKey64Bit
        
End Sub

' =============================================================================

'@Description("Deregisters a COM enumeration.")
'@Param("strGuid: The GUID of the enumeration to be deregistered.")
Public Static Sub DeregisterEnum(ByVal strGuid As String)

Dim strRecordKey32Bit As String
Dim strRecordKey64Bit As String
   
    ' SOFTWARE\CLASSES\Record\EnumGuid
    strRecordKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Record\" & strGuid
    mobjRegistry.DeleteRegistryTree strRecordKey32Bit

    strRecordKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Record\" & strGuid
    mobjRegistry.DeleteRegistryTree strRecordKey64Bit

End Sub

' =============================================================================

'@Description("Deregisters a COM interface.")
'@Param("strGuid: The GUID of the interface to be deregistered.")
Public Static Sub DeregisterInterface(ByVal strGuid As String)

Dim strInterfaceKey32Bit As String
Dim strInterfaceKey64Bit As String

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid
    strInterfaceKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Interface\" & strGuid
    mobjRegistry.DeleteRegistryTree strInterfaceKey32Bit
    
    strInterfaceKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Interface\" & strGuid
    mobjRegistry.DeleteRegistryTree strInterfaceKey64Bit
    
End Sub

' =============================================================================

'@Description("Deregisters a COM type library.")
Public Static Sub DeregisterTypeLibrary()

Dim strTypeLibKey32Bit As String
Dim strTypeLibKey64Bit As String

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid
    strTypeLibKey32Bit = mstrSoftwareClassesKeyBase32Bit & "TypeLib\" & Me.TypeLibraryGuid
    mobjRegistry.DeleteRegistryTree strTypeLibKey32Bit
        
    strTypeLibKey64Bit = mstrSoftwareClassesKeyBase64Bit & "TypeLib\" & Me.TypeLibraryGuid
    mobjRegistry.DeleteRegistryTree strTypeLibKey64Bit
        
End Sub

' =============================================================================

'@Description("Deregisters a VBE add-in component.")
'@Param("strProgId: The ProgID of the VBE add-in to be deregistered.")
Public Static Sub DeregisterVbeAddin(ByVal strProgId As String)
    
Const strVbeAddinKey32Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\"
Const strVbeAddinKey64Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns64\"
   
Dim strAddinClassKey32Bit As String
Dim strAddinClassKey64Bit As String
   
    ' SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\ProgId
    strAddinClassKey32Bit = strVbeAddinKey32Bit & strProgId
    mobjRegistry.DeleteRegistryTree strAddinClassKey32Bit
    
    strAddinClassKey64Bit = strVbeAddinKey64Bit & strProgId
    mobjRegistry.DeleteRegistryTree strAddinClassKey64Bit
    
End Sub

' =============================================================================

'@Description("Registers a COM class.")
'@Param("strName: The name of the class to be registered.")
'@Param("strGuid: The GUID of the class to be registered.")
'@Param("strProgId: The ProgID of the class to be registered.")
Public Static Sub RegisterClass(ByVal strName As String, ByVal strGuid As String, ByVal strProgId As String)

Const strComComponentGuid As String = "{62C8FE65-4EBB-45e7-B440-6E39B2CDBF29}"

Dim strClassRegistrationKeyBase32Bit As String
Dim strClsidKey32Bit As String
Dim strImplementedCategoriesKey32Bit As String
Dim strInprocServer32Key32Bit As String
Dim strInprocServer32AssemblyVersionKey32Bit As String
Dim strClsidProgIdKey32Bit As String

Dim strClassRegistrationKeyBase64Bit As String
Dim strClsidKey64Bit As String
Dim strImplementedCategoriesKey64Bit As String
Dim strInprocServer32Key64Bit As String
Dim strInprocServer32AssemblyVersionKey64Bit As String
Dim strClsidProgIdKey64Bit As String

    ' SOFTWARE\CLASSES\ProgId
    strClassRegistrationKeyBase32Bit = mstrSoftwareClassesKeyBase32Bit & strProgId
    mobjRegistry.CreateRegistryKey strClassRegistrationKeyBase32Bit
    mobjRegistry.WriteRegistryString strClassRegistrationKeyBase32Bit, vbNullString, strName

    strClassRegistrationKeyBase64Bit = mstrSoftwareClassesKeyBase64Bit & strProgId
    mobjRegistry.CreateRegistryKey strClassRegistrationKeyBase64Bit
    mobjRegistry.WriteRegistryString strClassRegistrationKeyBase64Bit, vbNullString, strName
    

    ' SOFTWARE\CLASSES\ProgId\CLSID
    strClsidKey32Bit = strClassRegistrationKeyBase32Bit & "\CLSID"
    mobjRegistry.CreateRegistryKey strClsidKey32Bit
    mobjRegistry.WriteRegistryString strClsidKey32Bit, vbNullString, strGuid

    strClsidKey64Bit = strClassRegistrationKeyBase64Bit & "\CLSID"
    mobjRegistry.CreateRegistryKey strClsidKey64Bit
    mobjRegistry.WriteRegistryString strClsidKey64Bit, vbNullString, strGuid
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid
    strClsidKey32Bit = mstrSoftwareClassesKeyBase32Bit & "CLSID\" & strGuid
    mobjRegistry.CreateRegistryKey strClsidKey32Bit
    mobjRegistry.WriteRegistryString strClsidKey32Bit, vbNullString, strName
    
    strClsidKey64Bit = mstrSoftwareClassesKeyBase64Bit & "CLSID\" & strGuid
    mobjRegistry.CreateRegistryKey strClsidKey64Bit
    mobjRegistry.WriteRegistryString strClsidKey64Bit, vbNullString, strName
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\Implemented Categories
    strImplementedCategoriesKey32Bit = strClsidKey32Bit & "\Implemented Categories"
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey32Bit
    
    strImplementedCategoriesKey64Bit = strClsidKey64Bit & "\Implemented Categories"
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey64Bit
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\Implemented Categories\COM COMPONENT GUID
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey32Bit & "\" & strComComponentGuid
    mobjRegistry.CreateRegistryKey strImplementedCategoriesKey64Bit & "\" & strComComponentGuid
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\InprocServer32
    strInprocServer32Key32Bit = strClsidKey32Bit & "\InprocServer32"
    mobjRegistry.CreateRegistryKey strInprocServer32Key32Bit
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, vbNullString, "mscoree.dll"
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "RuntimeVersion", Me.RuntimeVersion
    mobjRegistry.WriteRegistryString strInprocServer32Key32Bit, "ThreadingModel", "Both"

    strInprocServer32Key64Bit = strClsidKey64Bit & "\InprocServer32"
    mobjRegistry.CreateRegistryKey strInprocServer32Key64Bit
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, vbNullString, "mscoree.dll"
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "RuntimeVersion", Me.RuntimeVersion
    mobjRegistry.WriteRegistryString strInprocServer32Key64Bit, "ThreadingModel", "Both"
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\InprocServer32\AssemblyVersion
    strInprocServer32AssemblyVersionKey32Bit = strInprocServer32Key32Bit & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strInprocServer32AssemblyVersionKey32Bit
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey32Bit, "RuntimeVersion", Me.RuntimeVersion
    
    strInprocServer32AssemblyVersionKey64Bit = strInprocServer32Key64Bit & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strInprocServer32AssemblyVersionKey64Bit
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "Class", strName
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strInprocServer32AssemblyVersionKey64Bit, "RuntimeVersion", Me.RuntimeVersion
    

    ' SOFTWARE\CLASSES\CLSID\ClassGuid\ProgId
    strClsidProgIdKey32Bit = strClsidKey32Bit & "\ProgId"
    mobjRegistry.CreateRegistryKey strClsidProgIdKey32Bit
    mobjRegistry.WriteRegistryString strClsidProgIdKey32Bit, vbNullString, strProgId
    
    strClsidProgIdKey64Bit = strClsidKey64Bit & "\ProgId"
    mobjRegistry.CreateRegistryKey strClsidProgIdKey64Bit
    mobjRegistry.WriteRegistryString strClsidProgIdKey64Bit, vbNullString, strProgId
    
End Sub

' =============================================================================

'@Description("Registers a COM enumeration.")
'@Param("strName: The name of the enumeration to be registered.")
'@Param("strGuid: The GUID of the enumeration to be registered.")
Public Static Sub RegisterEnum(ByVal strName As String, ByVal strGuid As String)

Dim strRecordKey As String

    ' SOFTWARE\CLASSES\Record\EnumGuid\AssemblyVersion
    strRecordKey = mstrSoftwareClassesKeyBase64Bit & "Record\" & strGuid & "\" & Me.AssemblyVersion
    mobjRegistry.CreateRegistryKey strRecordKey
    mobjRegistry.WriteRegistryString strRecordKey, "Assembly", Me.AssemblyName
    mobjRegistry.WriteRegistryString strRecordKey, "Class", strName
    mobjRegistry.WriteRegistryString strRecordKey, "CodeBase", Me.AssemblyCodebasePath & "\" & Me.AssemblyCodebaseFile
    mobjRegistry.WriteRegistryString strRecordKey, "RuntimeVersion", Me.RuntimeVersion
    
End Sub

' =============================================================================

'@Description("Registers a COM interface.")
'@Param("strName: The name of the interface to be registered.")
'@Param("strGuid: The GUID of the interface to be registered.")
Public Static Sub RegisterInterface(ByVal strName As String, ByVal strGuid As String)

Dim strInterfaceKey32Bit As String
Dim strProxyStubClsid32Key32Bit As String
Dim strTypeLibraryKey32Bit As String

Dim strInterfaceKey64Bit As String
Dim strProxyStubClsid32Key64Bit As String
Dim strTypeLibraryKey64Bit As String

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid
    strInterfaceKey32Bit = mstrSoftwareClassesKeyBase32Bit & "Interface\" & strGuid
    mobjRegistry.CreateRegistryKey strInterfaceKey32Bit
    mobjRegistry.WriteRegistryString strInterfaceKey32Bit, vbNullString, strName

    strInterfaceKey64Bit = mstrSoftwareClassesKeyBase64Bit & "Interface\" & strGuid
    mobjRegistry.CreateRegistryKey strInterfaceKey64Bit
    mobjRegistry.WriteRegistryString strInterfaceKey64Bit, vbNullString, strName


    ' SOFTWARE\CLASSES\Interface\InterfaceGuid\ProxyStubClsid32
    strProxyStubClsid32Key32Bit = strInterfaceKey32Bit & "\ProxyStubClsid32"
    mobjRegistry.CreateRegistryKey strProxyStubClsid32Key32Bit
    mobjRegistry.WriteRegistryString strProxyStubClsid32Key32Bit, vbNullString, Me.TypeLibraryGuid
    
    strProxyStubClsid32Key64Bit = strInterfaceKey64Bit & "\ProxyStubClsid32"
    mobjRegistry.CreateRegistryKey strProxyStubClsid32Key64Bit
    mobjRegistry.WriteRegistryString strProxyStubClsid32Key64Bit, vbNullString, Me.TypeLibraryGuid
    

    ' SOFTWARE\CLASSES\Interface\InterfaceGuid\TypeLib
    strTypeLibraryKey32Bit = strInterfaceKey32Bit & "\TypeLib"
    mobjRegistry.CreateRegistryKey strTypeLibraryKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibraryKey32Bit, vbNullString, strGuid
    mobjRegistry.WriteRegistryString strTypeLibraryKey32Bit, "Version", Me.TypeLibraryVersion
    
    strTypeLibraryKey64Bit = strInterfaceKey64Bit & "\TypeLib"
    mobjRegistry.CreateRegistryKey strTypeLibraryKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibraryKey64Bit, vbNullString, strGuid
    mobjRegistry.WriteRegistryString strTypeLibraryKey64Bit, "Version", Me.TypeLibraryVersion
    
End Sub

' =============================================================================

'@Description("Registers a COM type library.")
Public Static Sub RegisterTypeLibrary()

Dim strTypeLibKey32Bit As String
Dim strTypeLibWin32Key32Bit As String
Dim strTypeLibFlagsKey32Bit As String
Dim strTypeLibHelpdirKey32Bit As String

Dim strTypeLibKey64Bit As String
Dim strTypeLibWin32Key64Bit As String
Dim strTypeLibFlagsKey64Bit As String
Dim strTypeLibHelpdirKey64Bit As String

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion
    strTypeLibKey32Bit = mstrSoftwareClassesKeyBase32Bit & "TypeLib\" & Me.TypeLibraryGuid & "\" & Me.TypeLibraryVersion
    mobjRegistry.CreateRegistryKey strTypeLibKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibKey32Bit, vbNullString, Me.TypeLibraryName
    
    strTypeLibKey64Bit = mstrSoftwareClassesKeyBase64Bit & "TypeLib\" & Me.TypeLibraryGuid & "\" & Me.TypeLibraryVersion
    mobjRegistry.CreateRegistryKey strTypeLibKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibKey64Bit, vbNullString, Me.TypeLibraryName
        

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\0\win32
    strTypeLibWin32Key32Bit = strTypeLibKey32Bit & "\0\win32"
    mobjRegistry.CreateRegistryKey strTypeLibWin32Key32Bit
    mobjRegistry.WriteRegistryString strTypeLibWin32Key32Bit, vbNullString, Me.TypeLibraryFilePath & "\" & Me.TypeLibraryFileName

    strTypeLibWin32Key64Bit = strTypeLibKey64Bit & "\0\win32"
    mobjRegistry.CreateRegistryKey strTypeLibWin32Key64Bit
    mobjRegistry.WriteRegistryString strTypeLibWin32Key64Bit, vbNullString, Me.TypeLibraryFilePath & "\" & Me.TypeLibraryFileName
    

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\FLAGS
    strTypeLibFlagsKey32Bit = strTypeLibKey32Bit & "\FLAGS"
    mobjRegistry.CreateRegistryKey strTypeLibFlagsKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibFlagsKey32Bit, vbNullString, "0"

    strTypeLibFlagsKey64Bit = strTypeLibKey64Bit & "\FLAGS"
    mobjRegistry.CreateRegistryKey strTypeLibFlagsKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibFlagsKey64Bit, vbNullString, "0"
    

    ' SOFTWARE\CLASSES\TypeLib\TypeLibGuid\TypeLibVersion\HELPDIR
    strTypeLibHelpdirKey32Bit = strTypeLibKey32Bit & "\HELPDIR"
    mobjRegistry.CreateRegistryKey strTypeLibHelpdirKey32Bit
    mobjRegistry.WriteRegistryString strTypeLibHelpdirKey32Bit, vbNullString, Me.TypeLibraryFilePath
    
    strTypeLibHelpdirKey64Bit = strTypeLibKey64Bit & "\HELPDIR"
    mobjRegistry.CreateRegistryKey strTypeLibHelpdirKey64Bit
    mobjRegistry.WriteRegistryString strTypeLibHelpdirKey64Bit, vbNullString, Me.TypeLibraryFilePath
    
End Sub

' =============================================================================

'@Description("Registers a VBE add-in component.")
'@Param("strProgId: The ProgID of the VBE add-in to be registerd.")
'@Param("strDescription: A short description of the VBE add-in.")
'@Param("strFriendlyName: A friendly name for the VBE add-in.")
Public Static Sub RegisterVbeAddin(ByVal strProgId As String, ByVal strDescription As String, ByVal strFriendlyName As String)
    
Const strVbeAddinKey32Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\"
Const strVbeAddinKey64Bit As String = "SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns64\"
Const intLoadBehaviorAtStartup As Integer = 3

Dim strAddinClassKey32Bit As String
Dim strAddinClassKey64Bit As String

    ' SOFTWARE\Microsoft\VBA\VBE\6.0\AddIns\ProgId
    strAddinClassKey32Bit = strVbeAddinKey32Bit & strProgId
    mobjRegistry.CreateRegistryKey strAddinClassKey32Bit
    mobjRegistry.WriteRegistryString strAddinClassKey32Bit, "Description", strDescription
    mobjRegistry.WriteRegistryString strAddinClassKey32Bit, "FriendlyName", strFriendlyName
    mobjRegistry.WriteRegistryInteger strAddinClassKey32Bit, "LoadBehavior", intLoadBehaviorAtStartup
    
    strAddinClassKey64Bit = strVbeAddinKey64Bit & strProgId
    mobjRegistry.CreateRegistryKey strAddinClassKey64Bit
    mobjRegistry.WriteRegistryString strAddinClassKey64Bit, "Description", strDescription
    mobjRegistry.WriteRegistryString strAddinClassKey64Bit, "FriendlyName", strFriendlyName
    mobjRegistry.WriteRegistryInteger strAddinClassKey64Bit, "LoadBehavior", intLoadBehaviorAtStartup
    
End Sub

' =============================================================================
