'Created by Joshua Greenert 10/20/2019
'This program will display the fiscal month range based on the parameters
'The parameters required are the period and year.
'When using the parameter in the function for the period, ensure to subtract one.

Public Function getFiscalStart(period, year) As Date
'Variable for return array and variables for start date.
Dim startDate
period = period - 2

'Declare array for start days.
Dim startDaysArray
startDaysArray = Array(28, 35, 28, 28, 35, 28, 28, 35, 28, 28, 35, 28)

'Declare begin date.
startDate = CDate(#2/3/2019#)

'Get year of the start date to use as a reference.
Dim startYear
startYear = DatePart("yyyy", startDate)

'Check if the year entered matches the year of the start record.
'If the year is greater than the current year, update the year to date.
If year = startYear Then
    'Add the days in the array to the year through a for loop.
    For i = 0 To period
        startDate = DateAdd("D", startDaysArray(i), startDate)
    Next
    
    'set the item to the object and send it.
    getFiscalStart = startDate
Else
    'Set a variable to determine the number of years.
    Dim numOfYears
    numOfYears = (year - startYear)
    
    'Use a multidimensional array to increase the date's value for each year.
    For j = 1 To numOfYears
        For i = 0 To 11
            startDate = DateAdd("D", startDaysArray(i), startDate)
        Next
    Next
    
    'Use another loop to get date brought to current period.
    For i = 0 To period
        startDate = DateAdd("D", startDaysArray(i), startDate)
    Next
    
    'set the item to the object and return it.
    getFiscalStart = startDate
End If

End Function

'This function will provide the end date.
Public Function getFiscalEnd(period, year) As Date
'Variable for return array and variables for start date.
Dim endDate
period = period - 2

'Declare array for end days.
Dim endDaysArray
endDaysArray = Array(35, 28, 28, 35, 28, 28, 35, 28, 28, 35, 28, 28)

'Declare end date.
endDate = CDate(#3/2/2019#)

'Get year of the start date to use as a reference.
Dim startYear
startYear = DatePart("yyyy", endDate)

'Check if the year entered matches the year of the start record.
'If the year is greater than the current year, update the year to date.
If year = startYear Then
    'Add the days in the array to the year through a for loop.
    For i = 0 To period
        endDate = DateAdd("D", endDaysArray(i), endDate)
    Next
    getFiscalEnd = endDate
Else
    'Set a variable to determine the number of years.
    Dim numOfYears
    numOfYears = (year - startYear)

'Use a multidimensional array to increase the date's value for each year.
For j = 1 To numOfYears
    For i = 0 To 11
        endDate = DateAdd("D", endDaysArray(i), endDate)
    Next
Next

'Use another loop to get date brought to current period.
For i = 0 To period
    endDate = DateAdd("D", endDaysArray(i), endDate)
Next
'set the item to the object and return it.
getFiscalEnd = endDate
End If

End Function