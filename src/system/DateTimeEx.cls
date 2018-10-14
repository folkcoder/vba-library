VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTimeEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'@Description("Provides methods to extend the DateTime class.")
'@Folder("VBALibrary.System")
'@PredeclaredId(True)
'@Remarks("This class has its hidden VB_PredeclaredId attribute set to true in order to create a static instance of the class without needing to be instantiated.")

Option Explicit

'@Description("Determines the nearest date, past or future, matching the provided day of the week.")
'@Param("intDayOfWeek: The day of the week to find.")
'@Param("dtmStartDate: The starting date to test.")
'@Returns("The nearest date (in the past or in the future) matching the provided day of the week.")
Public Function GetNearestWeekdayToDate(ByVal intDayOfWeek As VbDayOfWeek, ByVal dtmStartDate As Date) As Date

Dim intLastDayDate As Date
Dim intNextDayDate As Date

    ' Today is day requested
    If Weekday(Date) = intDayOfWeek Then
        GetNearestWeekdayToDate = Date
    Else
        
        intLastDayDate = GetPreviousDate(intDayOfWeek, dtmStartDate, True)
        intNextDayDate = GetNextDate(intDayOfWeek, dtmStartDate, True)
    
        ' Nearest day requested is in the past
        If Abs(DateDiff("d", Date, intLastDayDate)) < Abs(DateDiff("d", Date, intNextDayDate)) Then
            GetNearestWeekdayToDate = intLastDayDate
        
        ' Nearest day requested is in the future
        Else
            GetNearestWeekdayToDate = intNextDayDate
    
        End If
        
    End If

End Function


'@Description("Determines the next date matching the provided day of the week.")
'@Param("intDayOfWeek: The day of the week to find.")
'@Param("dtmStartDate: The starting date to test (defaults to current date).")
'@Param("blnMatchOnProvidedDate: A value indicating whether the function should return a match if the provided start date falls on the requested day of the week.")
'@Returns("The next date matching the provided day of the week.")
Public Function GetNextDate(ByVal intDayOfWeek As VbDayOfWeek, ByVal dtmStartDate As Date, Optional ByVal blnMatchOnProvidedDate As Boolean = True) As Date

Dim intWeekdayOfStartDate As Integer
    
    intWeekdayOfStartDate = Weekday(Date)
    If blnMatchOnProvidedDate = False And intWeekdayOfStartDate = intDayOfWeek Then
        GetNextDate = dtmStartDate + 7
    Else
        GetNextDate = dtmStartDate + 7 - Weekday(intWeekdayOfStartDate + 7 - intDayOfWeek)
    End If
    
End Function


'@Description("Determines the previous date matching the provided day of the week.")
'@Param("intDayOfWeek: The day of the week to find.")
'@Param("dtmStartDate: The starting date to test (defaults to current date).")
'@Param("blnMatchOnProvidedDate: A value indicating whether the function should return a match if the provided start date falls on the requested day of the week.")
'@Returns("The previous date matching the provided day of the week.")
Public Function GetPreviousDate(ByVal intDayOfWeek As VbDayOfWeek, ByVal dtmStartDate As Date, Optional ByVal blnMatchOnProvidedDate As Boolean = True) As Date

Dim intWeekdayOfStartDate As Integer

    intWeekdayOfStartDate = Weekday(Date)
    If blnMatchOnProvidedDate = True And intWeekdayOfStartDate = intDayOfWeek Then
        GetPreviousDate = Date
    Else
        GetPreviousDate = dtmStartDate - Weekday(dtmStartDate - intDayOfWeek)
    End If
    
End Function
