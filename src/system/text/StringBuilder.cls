VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'@Description("Provides methods to efficiently concatenate strings.")
'@Folder("VBALibrary.System.Text")
'@Remarks("The VBA native string concatenation operator '&' scales poorly when handling numerous concatenation operations or with very large strings. This class significantly improves performance in those scenarios.")
Option Compare Binary
Option Explicit

Private Type TStringBuilder
    Pointer As Long
    Text As String
End Type

Private this As TStringBuilder


' =============================================================================
' CONSTRUCTOR AND DESTRUCTOR
' =============================================================================

Private Sub Class_Initialize()
    Call Clear
End Sub

Private Sub Class_Terminate()
    Call Clear
End Sub


' =============================================================================
' PUBLIC PROPERTIES
' =============================================================================

'@Description("Gets the built string.")
Public Property Get Text() As String
    Text = Left$(this.Text, this.Pointer)
End Property

'@Description("Gets the length of the built string.")
Public Property Get Length() As Long
    Length = this.Pointer
End Property


' =============================================================================
' PUBLIC METHODS
' =============================================================================

'@Description("Appends the provided string to the current StringBuilder text.")
'@Param("strStringToAppend: The string to be appended to the current StringBuilder instance.")
Public Sub Append(ByVal strStringToAppend As String)

Dim lngNewStringLength As Long

    lngNewStringLength = Len(strStringToAppend)
    Call EnsureCapacity(lngNewStringLength)
    
    ' Append the new string to the first position of the allocated text.
    Mid$(this.Text, this.Pointer + 1, lngNewStringLength) = strStringToAppend
    
    ' Adjust the new pointer position
    this.Pointer = this.Pointer + lngNewStringLength
    
End Sub

' =============================================================================

'@Description("Appends the supplied string and a Carriage-Return/Linefeed to the end of the StringBuilder.")
'@Param("strTextToAppend: The string to be appended to the current StringBuilder instance.")
Public Sub AppendLine(Optional ByVal strTextToAppend As String)
    
    If LenB(strTextToAppend) > 0 Then
        Me.Append strTextToAppend
    End If
    
    Me.Append vbCrLf
    
End Sub

' =============================================================================

'@Description("Clears the StringBuilder instance.")
Public Sub Clear()
    this.Pointer = 0
    this.Text = vbNullChar
End Sub


' =============================================================================
' PRIVATE METHODS
' =============================================================================

'@Description("Ensures that the current text buffer has sufficient capacity for a new string.")
'@Param("lngIncreaseSize: The size by which the capacity needs to accommodate.")
Private Sub EnsureCapacity(ByVal lngIncreaseSize As Long)

Dim lngNewCapacity As Long
Dim lngRequiredCapacity As Long

    lngNewCapacity = this.Pointer * 2
    lngRequiredCapacity = lngIncreaseSize + this.Pointer
    
    ' New length required is more than the current length of the text
    If lngRequiredCapacity >= Len(this.Text) Then
    
        '  Length required is more than the standard increment
        If lngNewCapacity >= lngRequiredCapacity Then
            this.Text = this.Text & Space$(lngNewCapacity)
        Else
            this.Text = this.Text & Space$(lngRequiredCapacity)
        End If
    
    End If
End Sub

' =============================================================================
