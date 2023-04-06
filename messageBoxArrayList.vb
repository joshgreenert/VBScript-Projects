'Created by Joshua Greenert 10/24/2019
'This function will use the value entered to return a list of all items
'in each cell thereafter. Thus, performing the query function that is
'capable in Google Sheets. I hate you Excel you dirty, dirty girl.

Private Sub getListItems()
'Create an Arraylist object to add items to and variable to increment the number for row reference.
Dim arrList As ArrayList
Set arrList = New ArrayList
Dim issue
issue = Sheets("Overview").Range("E6")

Dim arrayAMT

'Set the column to check for instances for the issue, and the comments reported for it.
Dim issueColumn As Range
Set issueColumn = Sheets("Data").Range("E:E")
Dim commentColumn As Range
Set commentColumn = Sheets("Data").Range("F:F")

'Determine the number of non-empty cells in the column and set to object.
Dim columnLength
columnLength = WorksheetFunction.CountIf(Sheets("Data").Range("E:E"), "<>" & "")

Dim test

'For loop to add in every item into the list.
For i = 1 To columnLength
    test = issueColumn(i)
    If test = issue Then
        arrList.Add commentColumn(i)
    End If
Next

'Set item to size of arraylist - 1 to not go beyond index.
arrayAMT = arrList.Count - 1

'Set arrayList into array; Redim is needed for constant expression issue.
ReDim finalArray(arrayAMT)

'Place arrayList into array to use msgbox.
For i = 0 To arrayAMT
    finalArray(i) = arrList(i)
Next

'Send the message in a msgbox element.
MsgBox Join(finalArray, vbCrLf)
End Sub