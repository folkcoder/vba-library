VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder("VBALibrary.System.Text")
'@Description("Provides methods to extend the string class.")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")

Option Explicit

Private Declare PtrSafe Function lstrlenW Lib "kernel32" (ByVal lpString As LongPtr) As Long
Private Declare PtrSafe Function StrTrimW Lib "shlwapi" (ByVal pszSource As LongPtr, ByVal pszTrimChars As LongPtr) As Long


'@Description("Calculates the number of bytes in the provided string.")
'@Param("strStringToCalculate: The string to be calculated.")
'@Returns("The number of bytes in the string.")
Public Function CalculateBytes(ByVal strStringToCalculate As String) As Long
    CalculateBytes = LenB(strStringToCalculate) / 2
End Function


'@Description("Counts the number of times the provided text is found in the pro")
'@Param("strTextToBeSearched: The text to be searched.")
'@Param("strSearchTerm: The text for which to search.")
'@Returns("An lngeger representing the number of times a specified string occurrences in the provided search string.")
Public Function CountMatches(ByVal strTextToBeSearched As String, ByVal strSearchTerm As String) As Long
    CountMatches = Len(strTextToBeSearched) - Len(Replace(strTextToBeSearched, strSearchTerm, vbNullString))
End Function


'@Description("Generates a pseudo-random string.")
'@Param("intLength: The lenght of the string to be generated.")
'@Returns("A pseudo-random string.")
Public Function GenerateRandom(ByVal lngLength As Long) As String

Dim lngCharacterCount As Long
Dim lngRandomNumber As Long
Dim strResult As String

        Randomize

        For lngCharacterCount = 1 To lngLength
            lngRandomNumber = Int((126 - 33 + 1) * rnd + 33)
            Select Case lngRandomNumber
                Case 33, 35 To 38, 42, 48 To 57, 63 To 90, 97 To 122 ' !#$%^&* 0-9 ?@ A-Z a-z
                    strResult = strResult & Chr$(lngRandomNumber)
                Case Else
                    lngCharacterCount = lngCharacterCount - 1
            End Select
        Next lngCharacterCount
        
        GenerateRandom = strResult

End Function


'@Description("Determines whether the provided string matches a value in the provided delimited list.")
'@Param("strToEvaluate: The string to test.")
'@Param("strDelimitedList: A delimited list containing the strings to test against.")
'@Param("blnCaseSensitive: A value indicating whether the search should be case sensitive. Default is false.")
'@Param("strDelimiter: A value indicating the delimiter used in the delimited list. Default is the pipe character |.")
'@Returns("A value indicating whether the provided string is contained in the provided delimited list.")
Public Function InList(ByVal strToEvaluate As String, ByVal strDelimitedList As String, Optional ByVal blnCaseSensitive As Boolean = False, Optional ByVal strDelimiter As String = "|") As Boolean

Dim lngComparisonType As VbCompareMethod
Dim strValues() As String
Dim lngIndex As Long

    If blnCaseSensitive Then
        lngComparisonType = vbBinaryCompare
    Else
        lngComparisonType = vbTextCompare
    End If

    strValues = Split(strDelimitedList, strDelimiter, -1, lngComparisonType)
    For lngIndex = LBound(strValues) To UBound(strValues)
        If StrComp(strToEvaluate, strValues(lngIndex), lngComparisonType) = 0 Then
            InList = True
            Exit Function
        End If
    Next lngIndex
      
    InList = False

End Function


'@Description("Pads the provided string on the left side with the filler characters until the string's length matches the desired length.")
'@Param("strSource: The string to be padded.")
'@Param("intLength: The final length of the padded string.")
'@Param("strFillerCharacter: The character to be used to pad the string. Default is an empty space. Only single characters are permitted; additional characters will be ignored.")
'@Returns("A new string padded on the left with filler characters.")
Public Function PadLeft(ByVal strSource As String, ByVal lngLength As Long, Optional ByVal strFillerCharacter As String = " ") As String

    If Len(strSource) >= lngLength Then
        PadLeft = strSource
    Else
        PadLeft = Right$(String(lngLength, strFillerCharacter) & strSource, lngLength)
    End If

End Function


'@Description("Pads the provided string on the right side with the filler characters until the string's length matches the desired length.")
'@Param("strSource: The string to be padded.")
'@Param("intLength: The final length of the padded string.")
'@Param("strFillerCharacter: The character to be used to pad the string. Default is an empty space. Only single characters are permitted; additional characters will be ignored.")
'@Returns("A new string padded on the right with filler characters.")
Public Function PadRight(ByVal strSource As String, ByVal lngLength As Long, Optional ByVal strFillerCharacter As String = " ") As String

    If Len(strSource) >= lngLength Then
        PadRight = strSource
    Else
        PadRight = Left$(strSource & String(lngLength, strFillerCharacter), lngLength)
    End If

End Function


'@Description("Finds and replaces all instances of the provided search text except the first.")
'@Param("strSource: The string to be searched.")
'@Param("strTextToReplace: The string to be replaced.")
'@Param("strReplacementText: The string that will replace found matches.")
'@Returns("A string with all occurrences of the specified text replaced except the first.")
Public Function ReplaceExceptFirst(ByVal strSource As String, ByVal strTextToReplace As String, ByVal strReplacementText As String) As String
    ReplaceExceptFirst = Replace(strSource, strTextToReplace, strReplacementText, 1, -1, vbTextCompare)
End Function


'@Description("Finds and replaces all instances of the provided search text except the last.")
'@Param("strSource: The string to be searched.")
'@Param("strTextToReplace: The string to be replaced.")
'@Param("strReplacementText: The string that will replace found matches.")
'@Returns("A string with all occurrences of the specified text replaced except the last.")
Public Function ReplaceExceptLast(ByVal strSource As String, ByVal strTextToReplace As String, ByVal strReplacementText As String) As String

Dim lngMatchesCount As Long

    lngMatchesCount = CountMatches(strSource, strTextToReplace)
    ReplaceExceptLast = Replace(strSource, strTextToReplace, strReplacementText, 1, lngMatchesCount - 1, vbTextCompare)

End Function


'@Description("Finds and replaces the last character of a string.")
'@Param("strSource: The source string.")
'@Param("strReplacementText: The string that will replace the last character.")
'@Returns("A string with all occurrences of the specified text replaced except the last.")
Public Function ReplaceLastCharacter(ByVal strSource As String, ByVal strReplacementCharacter As String) As String
    ReplaceLastCharacter = Left$(strSource, Len(strSource) - 1) & strReplacementCharacter
End Function


'@Description("Removes nonalphanumeric characters from a string.")
'@Param("strSource: The string to be transformed.")
'@Param("blnStripAsterisks: A value indicating whether asterisks should be removed.")
'@Param("blnStripHyphens: A value indicating whether hyphens should be removed.")
'@Param("blnStripPipes: A value indicating whether vertical pipes should be removed.")
'@Param("blnStripSpaces: A value indicating whether spaces should be removed.")
'@Param("blnStripUnderscores: A value indicating underscores should be removed.")
'@Returns("A string containing the original string stripped of any nonalphanumeric characters.")
Public Function StripNonalphanumericChars(ByVal strSource As String, Optional ByVal blnStripAsterisks As Boolean = True, Optional ByVal blnStripHyphens As Boolean = False, Optional ByVal blnStripPipes As Boolean = True, Optional ByVal blnStripSpaces As Boolean = False, Optional ByVal blnStripUnderscores As Boolean = True) As String

Dim strResult As String
Dim lngCharacterCount As Long

    For lngCharacterCount = 1 To Len(strSource)
        Select Case Asc(Mid$(strSource, lngCharacterCount, 1))
            Case 48 To 57, 65 To 90, 97 To 122 ' 0-9 A-Z a-z
                strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
            Case 32 ' space
                If blnStripSpaces = False Then strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
            Case 42 ' *
                If blnStripAsterisks = False Then strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
            Case 45 ' -
                If blnStripHyphens = False Then strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
            Case 95
                If blnStripUnderscores = False Then strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
            Case 124
                If blnStripPipes = False Then strResult = strResult & Mid$(strSource, lngCharacterCount, 1)
        End Select
    Next
    StripNonalphanumericChars = strResult
    
End Function


'@Description("Trims all whitespace characters from a string rather than just spaces.")
'@Param("str: The string to be trimmed.")
'@Returns("A string trimmed of whitespace characters.")
Public Function TrimEx(ByVal str As String) As String

Const WHITE_SPACE As String = " " & vbTab & vbCr & vbLf

    If StrTrimW(StrPtr(str), StrPtr(WHITE_SPACE)) Then
        TrimEx = Left$(str, lstrlenW(StrPtr(str)))
    Else
        TrimEx = str
    End If

End Function
