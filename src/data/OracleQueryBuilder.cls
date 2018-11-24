VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OracleQueryBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Folder("VBALibrary.Data")
'@Description("Provides tools to construct Oracle PL/SQL statements containing predicates and parameters.")
'@Remarks("ODBC does not support named parameters. The class circumvents this limitation by determining parameter order within the query and creating an ordered collection of ADODB.Parameter objects.")

Option Explicit

Private Const mlngErrorNumber As Long = vbObjectError + 513
Private Const mstrClassName As String = "OracleQueryBuilder"
Private Const mstrParameterExistsErrorMessage As String = "A parameter with this name has already been added to the Parameters dictionary."

Private Type TQueryBuilder
    Parameters As Object
    QueryBody As String
    QueryFooter As String
End Type

Private mobjPredicates As Collection
Private this As TQueryBuilder


' =============================================================================
' CONSTRUCTOR / DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    Set this.Parameters = CreateObject("Scripting.Dictionary")
    Set mobjPredicates = New Collection
End Sub


' =============================================================================
' PROPERTIES
' =============================================================================

'@Description("Gets the collection of parameters.")
Public Property Get Parameters() As Collection
    Set Parameters = GetOrderedParameters
End Property

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
Dim param As Object
    If this.Parameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddBooleanParameter", mstrParameterExistsErrorMessage
    Else
        Set param = CreateObject("ADODB.Parameter")
        param.Name = strName
        param.Type = adBoolean
        param.Value = blnValue
        this.Parameters.Add strName, param
    End If
End Sub


'@Description("Maps a currency parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("curValue: The parameter's value.")
Public Sub AddCurrencyParameter(ByVal strName As String, ByVal curValue As Currency)
Dim param As Object
    If this.Parameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddCurrencyParameter", mstrParameterExistsErrorMessage
    Else
        Set param = CreateObject("ADODB.Parameter")
        param.Name = strName
        param.Type = adCurrency
        param.Value = curValue
        this.Parameters.Add strName, param
    End If
End Sub


'@Description("Maps a date parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("dtmValue: The parameter's value.")
Public Sub AddDateParameter(ByVal strName As String, ByVal dtmValue As Date)
Dim param As Object
    If this.Parameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddDateParameter", mstrParameterExistsErrorMessage
    Else
        Set param = CreateObject("ADODB.Parameter")
        param.Name = strName
        param.Type = adDBDate
        param.Value = dtmValue
        this.Parameters.Add strName, param
    End If
End Sub


'@Description("Maps a long parameter and its value to the query builder.")
'@Param("strName: The parameter's name.")
'@Param("lngValue: The parameter's value.")
Public Sub AddLongParameter(ByVal strName As String, ByVal lngValue As Long)
Dim param As Object
    If this.Parameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddNumericParameter", mstrParameterExistsErrorMessage
    Else
        Set param = CreateObject("ADODB.Parameter")
        param.Name = strName
        param.Type = adInteger
        param.Value = lngValue
        this.Parameters.Add strName, param
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
Dim param As Object
    If this.Parameters.Exists(strName) Then
        Err.Raise mlngErrorNumber, mstrClassName & ".AddStringParameter", mstrParameterExistsErrorMessage
    Else
        Set param = CreateObject("ADODB.Parameter")
        param.Name = strName
        param.Size = 255
        param.Type = adVarChar
        param.Value = strValue
        this.Parameters.Add strName, param
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


'@Description("Creates a new collection of ADODB.Parameter objects mapped to parameter placeholders in the query text and the class's parameters collection.")
'@Returns("A collection of ADODB.Parameter objects.")
Private Function GetOrderedParameters() As Collection

Dim colOrderedParameters As Collection
Dim objMatchedParameter As Object
Dim objNewParameter As Object

Dim strQuery As String
Dim objRegEx As Object
Dim objMatches As Object
Dim strParameterName As String

Dim lngCounter As Long

    Set colOrderedParameters = New Collection
    strQuery = this.QueryBody & " " & GetPredicatesText & " " & this.QueryFooter
    
    Set objRegEx = CreateObject("VBScript.RegExp")
    With objRegEx
        .Global = True
        .IgnoreCase = True
        .Multiline = True
        .Pattern = "@(\w+)"
        
        If objRegEx.test(strQuery) Then
            Set objMatches = objRegEx.Execute(strQuery)
            For lngCounter = 0 To objMatches.Count - 1
                
                strParameterName = objMatches(lngCounter).SubMatches(0)
                
                Set objMatchedParameter = this.Parameters(strParameterName)
                Set objNewParameter = CreateObject("ADODB.Parameter")
                With objNewParameter
                    .Name = lngCounter
                    .Size = objMatchedParameter.Size
                    .Type = objMatchedParameter.Type
                    .Value = objMatchedParameter.Value
                End With
                
                colOrderedParameters.Add objNewParameter
            Next lngCounter
        End If
    
    End With
    Set GetOrderedParameters = colOrderedParameters
    
End Function


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


'@Description("Replaces parameters in the predicates statements with a placeholder (?) for a value to be provided by a parameter object.")
'@Param("strPredicates: The text of the query's predicates.")
'@Returns("A string containing the predicates text with its parameters replaces by a parameter placeholder.")
Private Function ReplaceParametersWithValues(ByVal strPredicates As String) As String
    
Dim vntKey As Variant
Dim strParameterName As String
Dim strPredicatesWithValues As String

Const strProcedureName As String = "ReplaceParametersWithValues"

    strPredicatesWithValues = strPredicates
    For Each vntKey In this.Parameters.Keys
        strParameterName = CStr(vntKey)
       
        If InStr(1, strPredicatesWithValues, "@" & strParameterName, vbTextCompare) = 0 Then
            Err.Raise mlngErrorNumber, mstrClassName & "." & strProcedureName, "Parameter " & strParameterName & " was not found in the query."
        Else
            strPredicatesWithValues = Replace(strPredicatesWithValues, "@" & strParameterName, "?", 1, -1, vbTextCompare)
        End If
        
    Next vntKey
    
    ReplaceParametersWithValues = strPredicatesWithValues
    
End Function
