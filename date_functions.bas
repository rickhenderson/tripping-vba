Option Explicit

Private Function convert_to_date(stringDate As String) As Date
    '"""
    Since dates entered in the format "dd/mm/yyyy" are in an ambiguous format, 
    the convert_to_date function converts that one format into a standard Excel date value.
    '"""
    
    ' If a date value has been entered as "01/01/1961" which is "dd/mm/yyyy" then
    ' it is not a real date value in Excel. This makes extracting
    ' the day or month very difficult.
    
    ' This function accepts a date as an input argument and
    ' returns a valid date object.
    
    ' This will currently only work for dates written in "dd/mm/yyyy" format
    ' as described in the initial problem.
    
    Dim theMonth As Byte
    Dim theDay As Byte
    Dim theYear As Integer
    Dim currentDate As Date
    Dim realDate As String
    
    ' Get the month from the string date an convert it to a Byte value
    ' Byte is used to take up less memory
    theDay = CByte(Left(stringDate, 2))
    
    ' Since the format must be "dd/mm/yyyy" I can extract the value directly
    theMonth = CByte(Mid(stringDate, 4, 2)) ' Strings start at index 1, not 0
        
    theYear = CInt(Right(stringDate, 4))
    
    ' Create a specific date string to use the DateValue function
    ' to make it return the proper date. Otherwise strange things happen.
    realDate = theDay & "/" & theMonth & "/" & theYear
    
    ' Use DateValue to convert the string to a real date.
    currentDate = DateValue(realDate)
    
    ' Return the properly formatted Date as output from the function
    convert_to_date = currentDate
End Function
