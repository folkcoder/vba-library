VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MSAccessQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.Data")
'@Description("Provides tools to construct Microsoft Access SQL statements containing predicates and parameters.")

Option Explicit

Private Const mlngErrorNumber As Long = vbObjectError + 513
Private Const mstrClassName As String = "MSAccessQueryBuilder"
Private Const mstrParameterExistsErrorMessage As String = "A parameter with this name has already been added to the Parameters dictionary."

Private Type TSqlBuilder
    QueryBody As String
    QueryFooter As String
End Type

Private mobjParameters As Object
Private mobjPredicates As Collection
Private this As TSqlBuilder


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    Set mobjParameters = CreateObject("Scripting.Dictionary")
    Set mobjPredicates = New Collection
End Sub


' =============================================================================
' PROPERTIES
' =============================================================================

'@Description("Gets or sets the query statement (SELECT, INSERT, UPDATE, DELETE), exclusive of any predicates.")
Public Property Get QueryBody() As String
    QueryBody = this.QueryBody
End Property
Public Property Let QueryBody(ByVal Value As String)
    this.QueryBody = Value
End Property

'@Description("Gets or sets post-predicate query statements (e.g., GROUP BY, ORDER BY).")
Public Property Get QueryFooter() As String
    QueryFooter = this.QueryFooter
End Property
Public Property Let QueryFooter(ByVal Value As String)
    this.QueryFooter = Value
End Property


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Maps a boolean parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("blnValue: The parameter's value.")
Public Sub AddBooleanParameter(ByVal strName As String, ByVal blnValue As Boolean)
    If mobjParameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddBooleanParameter", mstrParameterExistsErrorMessage
    Else
        mobjParameters.Add strName, CStr(blnValue)
    End If
End Sub


'@Description("Maps a currency parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("curValue: The parameter's value.")
Public Sub AddCurrencyParameter(ByVal strName As String, ByVal curValue As Currency)
    If mobjParameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddCurrencyParameter", mstrParameterExistsErrorMessage
    Else
        mobjParameters.Add strName, CStr(curValue)
    End If
End Sub


'@Description("Maps a date parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("dtmValue: The parameter's value.")
Public Sub AddDateParameter(ByVal strName As String, ByVal dtmValue As Date)
    If mobjParameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddDateParameter", mstrParameterExistsErrorMessage
    Else
        mobjParameters.Add strName, "#" & CStr(dtmValue) & "#"
    End If
End Sub


'@Description("Maps a long parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("lngValue: The parameter's value.")
Public Sub AddLongParameter(ByVal strName As String, ByVal lngValue As Long)
    If mobjParameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddNumericParameter", mstrParameterExistsErrorMessage
    Else
        mobjParameters.Add strName, CStr(lngValue)
    End If
End Sub


'@Description("Adds a predicate to the query's WHERE criteria.")
'@Param("strPredicate: The predicate text to be added.")
Public Sub AddPredicate(ByVal strPredicate As String)
    mobjPredicates.Add "(" & strPredicate & ")"
End Sub


'@Description("Maps a string parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("strValue: The parameter's value.")
Public Sub AddStringParameter(ByVal strName As String, ByVal strValue As String)
    If mobjParameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddStringParameter", mstrParameterExistsErrorMessage
    Else
        mobjParameters.Add strName, "'" & strValue & "'"
    End If
End Sub


'@Description("Parses the query, its predicates, and any parameter values, and outputs an SQL statement.")
'@Returns("A string containing the parsed query.")
Public Function ToString() As String

Dim strPredicatesWithValues As String
    
    Const strErrorSource As String = "QueryBuilder.ToString"
    
    If this.QueryBody = vbNullString Then
        Err.Raise mlngErrorNumber, strErrorSource, "No query body is currently defined. Unable to build valid SQL."
    End If
    ToString = this.QueryBody
    
    strPredicatesWithValues = ReplaceParametersWithValues(GetPredicatesText)
    EnsureParametersHaveValues strPredicatesWithValues
    
    If Not strPredicatesWithValues = vbNullString Then
        ToString = ToString & " " & strPredicatesWithValues
    End If
    
    If Not this.QueryFooter = vbNullString Then
        ToString = ToString & " " & this.QueryFooter & ";"
    End If
    
End Function


' =============================================================================
' PRIVATE METHODS
' =============================================================================

'@Description("Ensures that all parameters defined in the query have been provided a value.")
'@Param("strQueryText: The query text to verify.")
Private Sub EnsureParametersHaveValues(ByVal strQueryText As String)

Dim strUnmatchedParameter As String
Dim lngMatchedPoisition As Long
Dim lngWordEndPosition As Long

    Const strProcedureName As String = "EnsureParametersHaveValues"

    lngMatchedPoisition = InStr(1, strQueryText, "@", vbTextCompare)
    If lngMatchedPoisition <> 0 Then
        lngWordEndPosition = InStr(lngMatchedPoisition, strQueryText, Space$(1), vbTextCompare)
        strUnmatchedParameter = Mid$(strQueryText, lngMatchedPoisition, lngWordEndPosition - lngMatchedPoisition)
    End If

    If Not strUnmatchedParameter = vbNullString Then
        Err.Raise mlngErrorNumber, mstrClassName & "." & strProcedureName, "Parameter " & strUnmatchedParameter & " has not been provided a value."
    End If
    
End Sub


'@Description("Combines each predicate in the predicates collection into a single string statement.")
'@Returns("A string containing the text of all predicates added to the query builder.")
Private Function GetPredicatesText() As String
    
Dim strPredicates As String
Dim vntPredicate As Variant

    If mobjPredicates.Count > 0 Then
        strPredicates = "WHERE 1 = 1"
        For Each vntPredicate In mobjPredicates
            strPredicates = strPredicates & " AND " & CStr(vntPredicate)
        Next vntPredicate
    End If
    
    GetPredicatesText = strPredicates
    
End Function


'@Description("Replaces parameters in the predicates statements with their provided values.")
'@Param("strPredicates: The text of the query's predicates.")
'@Returns("A string containing the predicates text with its parameters replaces by their provided values.")
Private Function ReplaceParametersWithValues(ByVal strPredicates As String) As String

Dim vntKey As Variant
Dim strParameterName As String
Dim strParameterValue As String
Dim strPredicatesWithValues As String

    Const strProcedureName As String = "ReplaceParametersWithValues"

    strPredicatesWithValues = strPredicates
    For Each vntKey In mobjParameters.Keys
        strParameterName = CStr(vntKey)
        strParameterValue = CStr(mobjParameters(vntKey))
        
        If InStr(1, strPredicatesWithValues, "@" & strParameterName, vbTextCompare) = 0 Then
            Err.Raise mlngErrorNumber, mstrClassName & "." & strProcedureName, "Parameter " & strParameterName & " was not found in the query."
        Else
            strPredicatesWithValues = Replace(strPredicatesWithValues, "@" & strParameterName, strParameterValue, 1, -1, vbTextCompare)
        End If
    Next vntKey
    
    ReplaceParametersWithValues = strPredicatesWithValues
    
End Function
