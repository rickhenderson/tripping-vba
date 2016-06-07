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

Option Explicit

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
Sub calculate_yearly_rainfall()

    Dim dataStartCell As Range
    Dim outputStartCell As Range
    Dim currentCell As Range
    Dim currentDate As String
    Dim currentMonth As String
    Dim currentYear As String
    Dim dataRange As Range
    Dim yearlySum As Double
    Dim previousMonth As String
    Dim previousYear As String
    Dim yearNum As Integer
    Dim monthList As String
    Dim outMonth As Byte
    Dim cell As Range
    
    Dim dateNum As Double
    Dim inputRangeSize As Integer
    
    ' NOTE: Need to correct one value from -1 to 0: Date 30/06/2013
    
    ' Turn screen updating off so the program runs faster
    ' This stops the screen from flickering as each value is written to the screen
    Application.ScreenUpdating = False
    
    Call ClearYearlyResults
    
    ' Create a range variable to point to the start of the data
    Set dataStartCell = Worksheets("Given Data Format").Range("A1")
           
    ' Create a range variable to point to where the output should start
    Set outputStartCell = Worksheets("Yearly Rainfall").Range("A1")
    
    ' Select all the date values and name the selected range
    With dataStartCell
        Range(.Offset(1, 0), .End(xlDown)).Name = "dataRange"
    End With
           
    ' Create a variable to easily refer to the data range
    Set dataRange = Range("dataRange")
    
    ' Set the yearNum to 1 to represent the first year of data for reporting purposes
    yearNum = 1
    
    ' Initialize the variables used as counters
    dateNum = 0
    yearlySum = 0
    currentMonth = getMonth(dataStartCell.Offset(1, 0).Value)
    previousMonth = getMonth(dataStartCell.Offset(1, 0).Value)
    previousYear = getYear(dataStartCell.Offset(1, 0).Value)
    
    ' Start the main loop
    With dataStartCell
    
        For Each cell In dataRange
            ' Go to the next row in the date set
            dateNum = dateNum + 1
            currentDate = cell.Value
            ' Get the current years and month - just string extraction
            currentYear = getYear(currentDate)
            currentMonth = getMonth(currentDate)
                
                ' if currentMonth = previousMonth then
                
                ' end if
                
                If currentYear = previousYear Then
                    ' Add the rainfall to the yearlySum
                    yearlySum = yearlySum + .Offset(dateNum, 1).Value
                Else
                    ' A new year has occurred
                    ' Output the current Year's rainfall
                    outputStartCell.Offset(yearNum, 0).Value = previousYear
                    outputStartCell.Offset(yearNum, 1).Value = yearlySum
                    yearNum = yearNum + 1
                    yearlySum = 0
                    previousYear = currentYear
                    'Debug.Print currentYear
                End If
        Next
        ' Need to output the row of actual data
        outputStartCell.Offset(yearNum, 0).Value = currentYear
        outputStartCell.Offset(yearNum, 1).Value = yearlySum
      
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub

Sub calculate_rainfall_by_month()

    Dim dataStartCell As Range
    Dim outputStartCell As Range
    Dim currentCell As Range
    Dim currentDate As String
    Dim currentMonth As String
    Dim currentYear As String
    Dim currentRainfall As Double
    Dim dataRange As Range
    Dim yearlySum As Double
    Dim monthlySum As Double
    Dim previousMonth As String
    Dim previousYear As String
    Dim monthList As String
    Dim monthNum As Byte
    Dim yearNum As Byte
    
    Dim cell As Range

    
    ' NOTE: You may need to correct one value from -1 to 0: Date 30/06/2013
    
    ' Turn screen updating off so the program runs faster
    ' This stops the screen from flickering as each value is written to the screen
    Application.ScreenUpdating = False
    
    ' Clear the previous run of this sub by deleting all output
    ' using a user defined function (UDF)
    'Call ClearMonthlyResults
    
    ' Create a range variable to point to the start of the data
    Set dataStartCell = Worksheets("Given Data Format").Range("A1")
           
    ' Create a range variable to point to where the output should start
    Set outputStartCell = Worksheets("Required Format").Range("A1")
    
    ' Select all the date values and name the selected range
    With dataStartCell
        Range(.Offset(1, 0), .End(xlDown)).Name = "dataRange"
    End With
           
    ' Create a variable to easily refer to the data range
    Set dataRange = Range("dataRange")
    
    ' Set the yearNum to 1 to represent the first year of data for reporting purposes
    yearNum = 1
    
    ' Initialize the variables used as counters

    yearlySum = 0
    monthlySum = 0
    monthNum = 1
    currentMonth = getMonth(dataStartCell.Offset(1, 0).Value)
    previousMonth = getMonth(dataStartCell.Offset(1, 0).Value)
    previousYear = getYear(dataStartCell.Offset(1, 0).Value)
    
    ' Start the main loop
    With dataStartCell
    
        For Each cell In dataRange
            ' Go to the next row in the date set
            'dateNum = dateNum + 1
            currentDate = cell.Value
            ' Get the current years and month - just string extraction
            currentYear = getYear(currentDate)
            currentMonth = getMonth(currentDate)
            
            ' Current rainfall is one column over from the current cell in the loop
            currentRainfall = cell.Offset(0, 1).Value
            
                If currentMonth = previousMonth Then
                    ' Add the rainfall to the current monthly sum
                    monthlySum = monthlySum + currentRainfall
                    
                Else
                    ' A new month has occurred
                    ' Output the month's total rainfall
                    outputStartCell.Offset(yearNum, monthNum).Value = monthlySum
                                       
                    ' If it was January, write the previous year in column A
                    If previousMonth = "01" Then
                        outputStartCell.Offset(yearNum, 0).Value = previousYear
                    End If
                                        
                    ' Increase the count to the next month
                    monthNum = monthNum + 1
                    
                    ' Set the previousMonth and previousYear to be the one just read
                    previousMonth = currentMonth
                    previousYear = currentYear
                    
                    ' Reset the monthly sum to 0
                    monthlySum = 0
                    
                End If
                
                If currentYear <> previousYear Then
                    ' A new year has occurred
                    
                    ' Add the December rainfall to the current monthly sum
                    monthlySum = monthlySum + currentRainfall
                    
                    ' Increase the count to the next month
                    monthNum = monthNum + 1
                    
                    ' Set the previousMonth as the currentMonth
                    previousMonth = currentMonth
                    Debug.Print ("December rainfall: " & monthlySum)
                    
                    ' Output the current Year's December Rainfall
                    outputStartCell.Offset(yearNum, monthNum).Value = monthlySum
                    
                    yearlySum = 0
                    yearNum = yearNum + 1
                    monthlySum = 0
                    monthNum = 1
                    previousYear = currentYear
                                        
                    
                    
                    'Debug.Print currentYear
                End If
               
        Next
        
    End With

    ' Turn screen updating back on
    Application.ScreenUpdating = True
End Sub


Sub ClearYearlyResults()
'
' ClearYearlyResults Macro
'

'
    Sheets("Yearly Rainfall").Select
    Range("A2:B2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("A2:B57").Select
    Selection.ClearContents
End Sub

Sub ClearMonthlyResults()
'
' ClearMonthlyResults Macro
'

'
    Range("A3:C3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents

End Sub
