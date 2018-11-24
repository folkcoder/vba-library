VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ActiveDirectoryViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.Network")
'@Description("Provides methods to query an Active Directory domain.")

Option Explicit

'@Description("Enumerates valid values for Active Directory attributes.")
Public Enum AttributeTypeEnum
    ADsPath
    Class
    CommonName
    Department
    EmailAddress
    EmployeeNumber
    EmployeeType
    GivenName
    ObjectName
    ObjectCategory
    PostalCode
    SamAccountName
    Surname
    State
    StreetAddress
    TelephoneNumber
    Title
End Enum

'@Description("Enumerates valid values for Active Directory extended attributes.")
Public Enum ExtendedAttributeTypeEnum
    Description
    Groups
    Members
    PublicKeys
    UID
End Enum

Private mstrLdapBase As String


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    mstrLdapBase = GetLDAPBase
End Sub


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Retrieves the value of the selected Active Directory attribute attribute for the provided SAMAccountName.")
'@Param("AttributeType: The name of the attribute to be retrieved.")
'@Param("strSamAccountName: The SAMAccountName of the Active Directory object to be queried (defaults to current user).")
'@Returns("An ActiveDirectory attribute for specified Active Directory object.").
Public Function GetAttribute(ByVal AttributeType As AttributeTypeEnum, Optional ByVal strSamAccountName As String = vbNullString) As String

Dim strDistinguishedName As String
Dim objAD As Object
    
    strDistinguishedName = GetDistinguishedname(strSamAccountName)
    Set objAD = GetObject("LDAP://" & strDistinguishedName)
    
    Select Case AttributeType
        Case AttributeTypeEnum.ADsPath
            GetAttribute = objAD.ADsPath
        
        Case AttributeTypeEnum.Class
            GetAttribute = objAD.Class
        
        Case AttributeTypeEnum.CommonName
            GetAttribute = objAD.cn
        
        Case AttributeTypeEnum.EmailAddress
            GetAttribute = objAD.Get("mail")
        
        Case AttributeTypeEnum.EmployeeNumber
            GetAttribute = objAD.Get("employeeNumber")
        
        Case AttributeTypeEnum.EmployeeType
            GetAttribute = objAD.Get("employeeType")
        
        Case AttributeTypeEnum.GivenName
            GetAttribute = objAD.Get("givenName")
        
        Case AttributeTypeEnum.ObjectName
            GetAttribute = objAD.Get("name")
        
        Case AttributeTypeEnum.ObjectCategory
            GetAttribute = objAD.Get("objectCategory")
        
        Case AttributeTypeEnum.PostalCode
            GetAttribute = objAD.Get("postalCode")
        
        Case AttributeTypeEnum.SamAccountName
            GetAttribute = objAD.Get("sAMAccountName")
        
        Case AttributeTypeEnum.Surname
            GetAttribute = objAD.Get("sn")
        
        Case AttributeTypeEnum.State
            GetAttribute = objAD.Get("st")
        
        Case AttributeTypeEnum.StreetAddress
            GetAttribute = objAD.Get("streetAddress")
        
        Case AttributeTypeEnum.TelephoneNumber
            GetAttribute = objAD.Get("telephoneNumber")
        
        Case AttributeTypeEnum.Title
            GetAttribute = objAD.Get("title")
            
    End Select
    
End Function


'@Description("Retrieves the distinguished name of the provided SAMAccountName.")
'@Param("strSamAccountName: The SAMAccountName of the Active Directory object to be queried (defaults to current user).")
'@Returns("A string containing a distinguished name.").
Public Function GetDistinguishedname(Optional ByVal strSamAccountName As String = vbNullString) As String
    
Dim strSearchName As String
Dim objConnection As Object
Dim objCommand As Object
Dim recLDAP As Object
        
    If strSamAccountName = vbNullString Then
        strSearchName = GetSamAccountName
    Else
        strSearchName = strSamAccountName
    End If
        
    Set objConnection = CreateObject("ADODB.Connection")
    Set objCommand = CreateObject("ADODB.Command")
    objConnection.Open "Provider=ADsDSOObject;"
    objCommand.ActiveConnection = objConnection
    
    objCommand.CommandText = "SELECT distinguishedName FROM " & mstrLdapBase & " WHERE sAMAccountName='" & strSearchName & "'"
    Set recLDAP = objCommand.Execute
    If recLDAP.RecordCount > 0 Then
        GetDistinguishedname = recLDAP.Fields(0).Value
    End If
        
End Function


'@Description("Retrieves the value of the selected Active Directory attribute attribute for the provided SAMAccountName.")
'@Param("ExtendedAttribute: The name of the extended attribute to be retrieved.")
'@Param("strSamAccountName: The SAMAccountName of the Active Directory object to be queried (defaults to current user).")
'@Returns("A multivalue attribute for the specified Active Directory object.")
Public Function GetExAttribute(ByVal ExtendedAttribute As ExtendedAttributeTypeEnum, Optional ByVal strSamAccountName As String = vbNullString) As Variant

Dim strDistinguishedName As String
Dim objAD As Object
    
    strDistinguishedName = GetDistinguishedname(strSamAccountName)
    Set objAD = GetObject("LDAP://" & strDistinguishedName)
    
    Select Case ExtendedAttribute
        
        Case ExtendedAttributeTypeEnum.Description
            GetExAttribute = objAD.GetEx("description")
        
        Case ExtendedAttributeTypeEnum.Groups ' returns distinguished name
            GetExAttribute = objAD.GetEx("memberOf")
        
        Case ExtendedAttributeTypeEnum.Members ' returns membership within group
            GetExAttribute = objAD.GetEx("member")
        
        Case ExtendedAttributeTypeEnum.PublicKeys ' returns hex
            GetExAttribute = objAD.GetEx("userCertificate")
        
        Case ExtendedAttributeTypeEnum.UID ' returns universal ID
            GetExAttribute = objAD.GetEx("uid")
            
    End Select
       
End Function


'@Description("Retrieves the organizational unit of the provided SAMAccountName.")
'@Param("strSamAccountName: The SAMAccountName of the Active Directory object to be queried (defaults to current user).")
'@Returns("A value from the second OU in a distinguished name.")
Public Function GetOrganizationalUnit(Optional ByVal strSamAccountName As String = vbNullString) As String

Dim strDistinguishedName As String
Dim vntDistinguishedNamePath As Variant
Dim intLength As Integer

    strDistinguishedName = GetDistinguishedname(strSamAccountName)
    vntDistinguishedNamePath = Split(strDistinguishedName, ",")
    intLength = Len(vntDistinguishedNamePath(3)) - 3
    GetOrganizationalUnit = Right$(vntDistinguishedNamePath(3), intLength)
            
End Function


'@Description("Retrieves the SAMAccountName of the current user.")
'@Returns("The SAMAccountName of the current user.")
Public Function GetSamAccountName() As String

Dim objSysInfo As Object
Dim strUser As String
Dim objUser As Object

    Set objSysInfo = CreateObject("ADSystemInfo")
    strUser = objSysInfo.UserName
    
    Set objUser = GetObject("LDAP://" & strUser)
    GetSamAccountName = objUser.SamAccountName

End Function


'@Description("Determines whether the provided SAMAccountName is a member of the provided group.")
'@Param("strGroupSamAccountName: The SAMAccountName of the group to be queried.")
'@Param("strUserSamAccountName: The SAMAccountName of the Active Directory object to be queried (defaults to current user).")
'@Returns("A value indicating whether the user is a member of the group.")
Public Function UserIsGroupMember(ByVal strGroupSamAccountName As String, Optional ByVal strUserSamAccountName As String = vbNullString) As Boolean

Dim strGroupDistinguishedName As String
Dim strUserDistinguishedName  As String
Dim objADUser As Object
Dim objADGroup As Object

    strGroupDistinguishedName = GetDistinguishedname(strGroupSamAccountName)
    strUserDistinguishedName = GetDistinguishedname(strUserSamAccountName)
    
    Set objADUser = GetObject("LDAP://" & strUserDistinguishedName)
    Set objADGroup = GetObject("LDAP://" & strGroupDistinguishedName)
    UserIsGroupMember = objADGroup.IsMember(objADUser.ADsPath)

End Function


' =============================================================================
' PRIVATE METHODS
' =============================================================================

'@Description("Retrieves the LDAP path of the current root domain.")
'@Returns("The LDAP path of the current root domain.")
Private Function GetLDAPBase() As String

Dim objADSysInfo As Object
Dim strADForest As String
        
    Set objADSysInfo = CreateObject("ADSystemInfo")
    strADForest = objADSysInfo.ForestDNSName
    GetLDAPBase = "'GC://" & strADForest & "'"
    
End Function
