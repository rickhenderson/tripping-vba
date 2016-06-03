Option Explicit

' ==============================================
' Given a worksheet with the following data:
' For example, containing daily rainfall readings every day for many years
'
' | Date         | Rainfall |
' | "01/01/2016" | 7.21     |
'      ...           ...
' This subroutine will calcluate the monthly totals of rainfall
' and output them to another worksheet by year and month in the following format:
'
' | Year | Jan | Feb | Mar |.........| 
' | 2016 | 34.4| 2.3 | 7.3 |.........|
'
' The sheet names are coded inside the sub so you will have to change those for re-use
' By using conditional formatting, this could potentially generate an Excel heatmap.

Sub calculate_montly_totals()
    Dim dataStartCell As Range
    Dim outputStartCell As Range
    Dim currentCell As Range
    Dim currentDate As String
    Dim currentMonth As String
    Dim currentYear As String
    Dim dataRange As Range
    Dim monthSum As Currency ' Currency might be more efficient than Double
    Dim previousMonth As String
    Dim previousYear As String
    Dim yearNum As Integer
    
    Dim r As Double
    Dim inputRangeSize As Integer
    
    ' Turn screen updating off so the program runs faster
    ' This stops the screen from flickering as each value is written to the screen
    Application.ScreenUpdating = False
    
    ' Create a range variable to point to the start of the data
    Set dataStartCell = Worksheets("Given Data Format").Range("A1")
        
    ' Set row counter to keep track of our position in the data list
    ' (a row offset)
    r = 1

    ' Create a range variable to point to where the output should start
    Set outputStartCell = Worksheets("Required Format").Range("A2")
    
    ' Select all the date values and name the selected range
    With dataStartCell
        Range(.Offset(1, 0), .End(xlDown)).Name = "dataRange"
    End With
           
    ' Create a variable to easily refer to the data range
    Set dataRange = Range("dataRange")
    
    ' Get the total size of the range
    inputRangeSize = dataRange.Rows.Count

    ' Set the yearNum to 1 to represent the first year of data for reporting purposes
    yearNum = 1
    
    ' Start the main loop
    With dataStartCell
       
        Do While r < inputRangeSize
            ' Calculate for the currentMonth
            ' Set the current cell to keep the code simple
            Set currentCell = dataStartCell.Offset(r)
             
            ' Get the date for the current data entry
            currentDate = currentCell.Value
            
            ' Extract the month & year from the date being currently read from the worksheet
            currentMonth = getMonth(currentDate)
            currentYear = getYear(currentDate)
            ' Debug.Print MonthName(currentMonth) # Neat feature to use if necessary
               
            ' Get the value from the rainfall column
            monthSum = monthSum + currentCell.Offset(0, 1).Value
            
            ' If the current month is different from the next month read in, and you aren't in the first row, reset the sum to 0
            If currentMonth <> previousMonth And r <> 1 Then
                ' Write out the monthSum for debugging or reporting
                'Debug.Print MonthName(currentMonth) & ", " & currentYear & ": " & monthSum
                
                ' Store the monthly result on the result worksheet in the right position
                With outputStartCell
                    .Offset(yearNum, currentMonth).Value = monthSum
                End With
                
                ' Reset the monthly sum
                monthSum = 0
                
                ' If the next entry falls in a new year, increase the year count for report purposes
                If currentYear <> previousYear Then
                    yearNum = yearNum + 1
                End If
            End If
            
         ' Move to the next row
         previousMonth = currentMonth
         previousYear = currentYear
         r = r + 1
        Loop
    
    End With
    
    ' Turn screen updating back on
    Application.ScreenUpdating = False
End Sub

Private Function getMonth(stringDate As String) As String
    ' Silly string extraction for string dates
    ' Since the format must be "dd/mm/yyyy" I can extract the value directly
    getMonth = Mid(stringDate, 4, 2) ' Strings start at index 1, not 0
End Function

Private Function getYear(stringDate As String) As String
 '   Return just the year
     getYear = Right(stringDate, 4)
End Function

Private Function getDay(stringDate As String) As String
 '   Return just the day
     getDay = Left(stringDate, 2)
End Function
